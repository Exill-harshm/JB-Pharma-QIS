from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ApiInfo:
    api_name: str
    manufacturer_text: str


@dataclass(frozen=True)
class SummaryInfo:
    summary_values_by_label: dict[str, list[str]]
    related_row_values: list[str]


@dataclass(frozen=True)
class ManufactureInfo:
    section_heading: str
    subtitle: str
    name_and_address: str
    responsibility: str
    api_pq_number: str
    letter_of_access: str


@dataclass(frozen=True)
class P31ManufacturerInfo:
    section_heading: str
    name_and_address: str
    responsibility: str


@dataclass(frozen=True)
class PipelineResult:
    output_docx: Path
    warnings: list[str]
