from __future__ import annotations

from pathlib import Path

import fitz

from .config import PipelineConfig
from .extractor import ApiInfoExtractor
from .filler import QisDocxFiller
from .models import PipelineResult
from .section_mapper import SectionMapper


class QisApiPipeline:
    def __init__(self, config: PipelineConfig):
        self._config = config
        self._mapper = SectionMapper(config.dossier_root)
        self._extractor = ApiInfoExtractor()
        self._filler = QisDocxFiller()

    def run(self) -> PipelineResult:
        warnings: list[str] = []
        self._config.artifacts_dir.mkdir(parents=True, exist_ok=True)

        api_pdf_path = self._mapper.resolve_pdf(self._config.ctd_reference)
        if api_pdf_path is None:
            raise FileNotFoundError(
                f"Could not resolve PDF for CTD section {self._config.ctd_reference} under {self._config.dossier_root}"
            )

        summary_pdf_path = self._resolve_module1_pdf()
        if summary_pdf_path is None:
            warnings.append("Could not identify a Module 1 PDF for QIS summary extraction.")

        p31_pdf_path = self._mapper.resolve_pdf("3.2.P.3.1")
        if p31_pdf_path is None:
            warnings.append("Could not resolve PDF for CTD section 3.2.P.3.1")

        api_info = self._extractor.extract(api_pdf_path)
        summary_info = self._extractor.extract_summary_info(summary_pdf_path) if summary_pdf_path else None
        manufacture_info = self._extractor.extract_manufacture_info(api_pdf_path)
        p31_info = self._extractor.extract_p31_manufacturer_info(p31_pdf_path) if p31_pdf_path else None
        if not api_info.api_name:
            warnings.append("Could not confidently extract API name.")
        if not api_info.manufacturer_text:
            warnings.append("Could not confidently extract API manufacturer text.")
        if not manufacture_info.subtitle:
            warnings.append("Could not confidently extract 2.3.S.2.1 manufacture details.")

        fill_warnings = self._filler.fill(
            self._config.template_docx,
            self._config.output_docx,
            api_info,
            summary_info,
            manufacture_info,
            p31_info,
        )
        warnings.extend(fill_warnings)
        self._write_log(api_pdf_path, summary_pdf_path, p31_pdf_path, api_info, summary_info, manufacture_info, p31_info)

        return PipelineResult(output_docx=self._config.output_docx, warnings=warnings)

    def _write_log(self, api_pdf_path: Path, summary_pdf_path: Path | None, p31_pdf_path: Path | None, api_info, summary_info, manufacture_info, p31_info) -> None:
        log_path = self._config.artifacts_dir / "qis_api_generation.log"
        lines = [
            f"Resolved API source PDF: {api_pdf_path}",
            f"Resolved summary PDF: {summary_pdf_path}",
            f"Resolved 3.2.P.3.1 PDF: {p31_pdf_path}",
            f"Extracted API name: {api_info.api_name}",
            "Extracted manufacturer text:",
            api_info.manufacturer_text,
            "",
            "Extracted manufacture section:",
            f"Heading: {manufacture_info.section_heading}",
            f"Subtitle: {manufacture_info.subtitle}",
            f"Manufacturer row: {manufacture_info.name_and_address}",
        ]

        if summary_info is not None:
            lines.extend(
                [
                    "",
                    "Extracted summary table values:",
                    f"Summary rows captured: {len(summary_info.summary_values_by_label)}",
                    f"Related dossier row captured: {bool(summary_info.related_row_values)}",
                ]
            )

        if p31_info is not None:
            lines.extend(
                [
                    "",
                    "Extracted 2.3.P.3.1 section:",
                    f"Heading: {p31_info.section_heading}",
                    f"Responsibility: {p31_info.responsibility}",
                ]
            )

        log_path.write_text("\n".join(lines), encoding="utf-8")

    def _resolve_module1_pdf(self) -> Path | None:
        all_pdfs = list(self._config.dossier_root.rglob("*.pdf"))
        if not all_pdfs:
            return None

        module1_candidates = [
            p
            for p in all_pdfs
            if any(part.lower().replace(" ", "") in {"module1", "m1"} for part in p.parts)
        ]
        if not module1_candidates:
            return None

        name_preferred = [
            p
            for p in module1_candidates
            if any(token in p.name.lower() for token in ["module 1", "module1", "qis", "quality information summary"])
        ]

        scan_pool = name_preferred or module1_candidates
        scan_pool = sorted(scan_pool, key=lambda p: p.stat().st_size, reverse=True)

        for candidate in scan_pool[:40]:
            if self._contains_qis_heading(candidate):
                return candidate

        return scan_pool[0]

    @staticmethod
    def _contains_qis_heading(pdf_path: Path) -> bool:
        needles = [
            "quality information summary (qis)",
            "1.4.2 the quality information summary",
        ]
        try:
            doc = fitz.open(pdf_path)
        except Exception:
            return False

        try:
            for i in range(min(doc.page_count, 12)):
                text = ApiInfoExtractor._read_page_text(doc.load_page(i)).lower()
                if any(needle in text for needle in needles):
                    return True
            return False
        finally:
            doc.close()
