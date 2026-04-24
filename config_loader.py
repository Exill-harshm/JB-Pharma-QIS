"""
Module: config_loader
Responsibility: Reads config.yaml, validates all paths, and returns a typed config object.
"""

import os
import yaml
from dataclasses import dataclass, field
from typing import Dict, Set


@dataclass
class Config:
    """Strongly typed configuration object"""
    template_docx_path:     str
    source_pdf_folder:      str
    output_docx_path:       str
    log_folder:             str
    enable_qis_v2_overlay:  bool          = True
    include_pdf_tables:     bool          = False
    dossier_root:           str           = ""
    mapping_logic_pdf_path: str           = ""
    section_page_limits:    Dict[str, int] = field(default_factory=dict)
    section_start_pages:    Dict[str, int] = field(default_factory=dict)
    table_only_sections:    Set[str]       = field(default_factory=set)
    table_only_all_sections: bool          = False
    table_keyword_by_template_section: Dict[str, str] = field(default_factory=dict)


def _as_bool(value, default: bool = True) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off"}:
        return False
    return default


def load_config(config_path: str = "config.yaml") -> Config:
    """
    Reads config.yaml, validates paths exist before starting,
    and returns a typed Config object.
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    with open(config_path, "r", encoding="utf-8") as file:
        try:
            data = yaml.safe_load(file)
        except yaml.YAMLError as e:
            raise ValueError(f"Failed to parse YAML configuration: {e}")

    required_keys = [
        "template_docx_path",
        "source_pdf_folder",
        "output_docx_path",
        "log_folder",
    ]
    for key in required_keys:
        if key not in data:
            raise KeyError(f"Missing required configuration key: {key}")

    if not os.path.exists(data["template_docx_path"]):
        raise FileNotFoundError(
            f"Template DOCX not found: {data['template_docx_path']}"
        )

    if not os.path.exists(data["source_pdf_folder"]):
        raise NotADirectoryError(
            f"Source PDF folder not found: {data['source_pdf_folder']}"
        )

    output_dir = os.path.dirname(data["output_docx_path"])
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    if not os.path.exists(data["log_folder"]):
        os.makedirs(data["log_folder"], exist_ok=True)

    raw_dossier_root = str(data.get("dossier_root", "") or "").strip()
    if raw_dossier_root and not os.path.exists(raw_dossier_root):
        raise NotADirectoryError(
            f"Configured dossier_root not found: {raw_dossier_root}"
        )

    enable_qis_v2_overlay = _as_bool(
        data.get("enable_qis_v2_overlay", True), default=True
    )
    include_pdf_tables = _as_bool(
        data.get("include_pdf_tables", False), default=False
    )

    raw_limits = data.get("section_page_limits", {}) or {}
    section_page_limits = {
        str(k): int(v) for k, v in raw_limits.items()
    }

    raw_starts = data.get("section_start_pages", {}) or {}
    section_start_pages = {
        str(k): int(v) for k, v in raw_starts.items()
    }

    raw_table_only = data.get("table_only_sections", []) or []
    table_only_sections = {str(s).strip() for s in raw_table_only if str(s).strip()}
    table_only_all_sections = bool(data.get("table_only_all_sections", False))
    raw_keyword_map = data.get("table_keyword_by_template_section", {}) or {}
    table_keyword_by_template_section = {
        str(k).strip(): str(v).strip()
        for k, v in raw_keyword_map.items()
        if str(k).strip() and str(v).strip()
    }

    return Config(
        template_docx_path     = data["template_docx_path"],
        mapping_logic_pdf_path = data.get("mapping_logic_pdf_path", ""),
        source_pdf_folder      = data["source_pdf_folder"],
        output_docx_path       = data["output_docx_path"],
        log_folder             = data["log_folder"],
        enable_qis_v2_overlay  = enable_qis_v2_overlay,
        include_pdf_tables     = include_pdf_tables,
        dossier_root           = raw_dossier_root,
        section_page_limits    = section_page_limits,
        section_start_pages    = section_start_pages,
        table_only_sections    = table_only_sections,
        table_only_all_sections = table_only_all_sections,
        table_keyword_by_template_section = table_keyword_by_template_section,
    )