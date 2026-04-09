from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class PipelineConfig:
    template_docx: Path
    dossier_root: Path
    output_docx: Path
    artifacts_dir: Path
    ctd_reference: str = "3.2.S.2.1"
