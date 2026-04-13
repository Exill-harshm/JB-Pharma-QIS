from __future__ import annotations

import argparse
from pathlib import Path

from src.qis_api.config import PipelineConfig
from src.qis_api.pipeline import QisApiPipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Fill QIS API box from section 3.2.S.2.1")
    parser.add_argument("--template", default="", help="Path to QIS template DOCX")
    parser.add_argument("--dossier-root", default="", help="Path to dossier root folder")
    parser.add_argument("--output", default="", help="Output DOCX path")
    parser.add_argument(
        "--artifacts-dir",
        default=str(Path(__file__).resolve().parent / "artifacts"),
        help="Artifacts folder for logs",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()

    required = [args.template, args.dossier_root, args.output]
    if not all(required):
        raise ValueError("Provide explicit --template --dossier-root --output")
    template = Path(args.template)
    dossier_root = Path(args.dossier_root)
    output = Path(args.output)

    config = PipelineConfig(
        template_docx=template,
        dossier_root=dossier_root,
        output_docx=output,
        artifacts_dir=Path(args.artifacts_dir),
    )

    result = QisApiPipeline(config).run()
    print("=== QIS API Generation Summary ===")
    print(f"Output DOCX: {result.output_docx}")
    print("Warnings:")
    for warning in result.warnings:
        print(f"- {warning}")


if __name__ == "__main__":
    main()
