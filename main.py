"""
Module: main
Responsibility: Orchestrates the full QIS generation pipeline.
config -> section map -> extract -> inject -> save DOCX -> summary
"""
import os
import sys

os.environ["PYTHONIOENCODING"] = "utf-8"
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from logger_setup import get_logger
from config_loader import load_config
from section_mapper import build_section_map
from docx_builder import process_template
from v2_overlay import apply_qis_v2_overlay


def main():
    try:
        config = load_config("config.yaml")
    except Exception as e:
        print(f"CRITICAL: Failed to load config.yaml: {e}")
        sys.exit(1)

    logger = get_logger(config.log_folder)
    logger.info("======== QIS DOCX Generation Pipeline Starting ========")

    try:
        section_map = build_section_map(
            source_folder    = config.source_pdf_folder,
            mapping_doc_path = config.mapping_logic_pdf_path,
            log_folder       = config.log_folder
        )
        logger.info(f"Mapped sections: {sorted(list(section_map.keys()))}")

        if not section_map:
            logger.error("No sections could be mapped. Aborting.")
            sys.exit(1)

        sections_filled, warnings, failures = process_template(
            template_path       = config.template_docx_path,
            output_path         = config.output_docx_path,
            section_map         = section_map,
            log_folder          = config.log_folder,
            section_page_limits = config.section_page_limits,
            section_start_pages = config.section_start_pages,
            preserve_template_tables = config.enable_qis_v2_overlay,
            include_pdf_tables  = config.include_pdf_tables,
            table_only_sections = config.table_only_sections,
            table_only_all_sections = config.table_only_all_sections,
            table_keyword_by_template_section = config.table_keyword_by_template_section,
        )

        v2_overlay_warnings = []
        if config.enable_qis_v2_overlay:
            logger.info("Starting QIS v2 overlay stage.")
            v2_overlay_warnings = apply_qis_v2_overlay(
                output_docx_path=config.output_docx_path,
                source_pdf_folder=config.source_pdf_folder,
                log_folder=config.log_folder,
                dossier_root=config.dossier_root,
            )
            for msg in v2_overlay_warnings:
                logger.warning(msg)
            if not v2_overlay_warnings:
                logger.info("QIS v2 overlay completed successfully.")

        warnings += len(v2_overlay_warnings)

        summary_lines = [
            "",
            "=" * 50,
            "      FINAL GENERATION SUMMARY",
            "=" * 50,
            f"  Sections successfully filled : {sections_filled}",
            f"  Warnings generated           : {warnings}",
            f"  Total failures               : {failures}",
            f"  QIS v2 overlay               : {'ENABLED' if config.enable_qis_v2_overlay else 'DISABLED'}",
            f"  Output DOCX                  : {config.output_docx_path}",
            "=" * 50,
        ]
        summary = "\n".join(summary_lines)
        try:
            print(summary)
        except UnicodeEncodeError:
            print(summary.encode('ascii', errors='replace').decode('ascii'))

        logger.info(summary)

    except Exception as e:
        logger.critical(f"Pipeline crashed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()