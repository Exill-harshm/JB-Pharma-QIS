"""
Module: v2_overlay
Responsibility: Applies QIS v2 table-driven overlay on top of an already
generated QIS DOCX so legacy section filling and v2 table filling both work.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _derive_dossier_root(source_pdf_folder: str) -> Path | None:
    """Best-effort dossier root derivation from the configured Module 3 folder."""
    source_path = Path(source_pdf_folder).resolve()

    for candidate in [source_path] + list(source_path.parents):
        if (candidate / "Module 1").exists() or (candidate / "m1").exists():
            return candidate

    compact_parts = [part.lower().replace(" ", "") for part in source_path.parts]
    for index, part in enumerate(compact_parts):
        if part in {"module3", "m3"} and index > 0:
            return Path(*source_path.parts[:index])

    return None


def apply_qis_v2_overlay(
    output_docx_path:  str,
    source_pdf_folder: str,
    log_folder:        str,
    dossier_root:      str = "",
) -> list[str]:
    """
    Runs the v2 QIS pipeline against an already generated DOCX.

    Returns warning messages (empty list means success).
    """
    warnings: list[str] = []

    output_docx = Path(output_docx_path)
    if not output_docx.exists():
        return [f"[QIS v2] Output DOCX not found for overlay: {output_docx}"]

    if dossier_root and dossier_root.strip():
        resolved_root = Path(dossier_root)
    else:
        resolved_root = _derive_dossier_root(source_pdf_folder) or Path()

    if not resolved_root.exists():
        return [
            "[QIS v2] Could not resolve dossier root. "
            "Set dossier_root in config.yaml to enable v2 overlay."
        ]

    repo_root = Path(__file__).resolve().parent
    v2_src = repo_root / "qis_v2" / "src"
    if not v2_src.exists():
        return [f"[QIS v2] Missing v2 source folder: {v2_src}"]

    run_script = repo_root / "qis_v2" / "run.py"
    if not run_script.exists():
        return [f"[QIS v2] Missing v2 runner: {run_script}"]

    artifacts_dir = Path(log_folder) / "qis_v2_artifacts"

    cmd = [
        sys.executable,
        str(run_script),
        "--template",
        str(output_docx),
        "--dossier-root",
        str(resolved_root),
        "--output",
        str(output_docx),
        "--artifacts-dir",
        str(artifacts_dir),
    ]

    completed = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).strip()
        return [f"[QIS v2] Overlay failed (exit {completed.returncode}): {detail}"]

    parsed_warnings: list[str] = []
    for line in completed.stdout.splitlines():
        striped = line.strip()
        if striped.startswith("-"):
            parsed_warnings.append(striped.lstrip("-").strip())

    suppress_prefixes = (
        "Could not locate 2.3.S.2.1 section in QIS template.",
        "Could not locate 2.3.P.3.1 section in QIS template.",
    )

    for warning in parsed_warnings:
        if warning.startswith(suppress_prefixes):
            continue
        warnings.append(f"[QIS v2] {warning}")
    return warnings
