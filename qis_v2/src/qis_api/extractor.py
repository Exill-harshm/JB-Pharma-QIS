from __future__ import annotations

import re
from pathlib import Path

import fitz

from .models import ApiInfo, ManufactureInfo, P31ManufacturerInfo, SummaryInfo


BODY_CLIP_TOP_RATIO = 0.10
BODY_CLIP_BOTTOM_RATIO = 0.92


class ApiInfoExtractor:
    def extract_summary_info(self, pdf_path: Path) -> SummaryInfo:
        doc = fitz.open(pdf_path)
        try:
            pages = [self._read_page_text(doc.load_page(i)) for i in range(doc.page_count)]
            q_page = self._find_summary_page_index(pages)
            if q_page is None:
                q_page = self._find_page_index(pages, "1.4.2 THE QUALITY INFORMATION SUMMARY (QIS)")
            if q_page is None:
                q_page = self._find_page_index(pages, "QUALITY INFORMATION SUMMARY (QIS)")
            if q_page is None:
                q_page = 0

            summary_values_by_label: dict[str, list[str]] = {}
            related_row_values: list[str] = []

            for page_index in range(q_page, min(q_page + 2, doc.page_count)):
                page = doc.load_page(page_index)
                clip_rect = self._body_clip_rect(page)
                tables = page.find_tables(clip=clip_rect).tables
                for table in tables:
                    rows = table.extract()
                    if not rows:
                        continue

                    related_candidate = self._extract_related_values_from_table(rows)
                    if related_candidate:
                        related_row_values = related_candidate
                        continue

                    page_map = self._rows_to_label_value_map(rows)
                    for key, values in page_map.items():
                        if key not in summary_values_by_label:
                            summary_values_by_label[key] = values

            self._repair_merged_address_rows(summary_values_by_label)
            self._repair_town_province_rows(summary_values_by_label)

            return SummaryInfo(
                summary_values_by_label=summary_values_by_label,
                related_row_values=related_row_values,
            )
        finally:
            doc.close()

    def extract_manufacture_info(self, pdf_path: Path) -> ManufactureInfo:
        text = self._read_text(pdf_path)
        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
        compact = re.sub(r"\s+", " ", text)

        api_name = self._extract_api_name(text)
        company = self._company_from_manufactured_by(lines)
        if not company:
            company = self._match(
                compact,
                [r"Name and address of API\(s\) Manufacturer\s+([^\n]+)"],
                "",
            )

        address = self._first_address_block(lines)
        if not address:
            address = self._match(
                compact,
                [r"(No\.[^\.]{20,250}\.)", r"(Plot\s*No\.[^\.]{20,250}\.)"],
                "",
            )

        subtitle = ""
        if api_name and company:
            subtitle = f"({api_name}, {company})"
        elif company:
            subtitle = f"({company})"

        name_and_address = company
        if address:
            if name_and_address:
                name_and_address += "\n\n"
            name_and_address += f"Address of Manufacturer:\n{address}"

        return ManufactureInfo(
            section_heading="2.3.S.2.1 Manufacturer(s)",
            subtitle=subtitle,
            name_and_address=name_and_address,
            responsibility="Manufacturing, Packaging and\nTesting",
            api_pq_number="Not applicable",
            letter_of_access="NO",
        )

    def extract_p31_manufacturer_info(self, pdf_path: Path) -> P31ManufacturerInfo:
        text = self._read_text(pdf_path)
        lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
        compact = re.sub(r"\s+", " ", text)

        registered = self._block_between(
            lines,
            start_tokens=["Address of Registered Office"],
            end_tokens=["Address of the Manufacturing Site", "Address of Manufacturing Site", "Certificate"],
        )
        manufacturing = self._block_between(
            lines,
            start_tokens=["Address of the Manufacturing Site", "Address of Manufacturing Site"],
            end_tokens=["Certificate", "2 of", "1 of"],
        )

        if not registered and not manufacturing:
            name_addr_block = self._block_between(
                lines,
                start_tokens=["Name and Address"],
                end_tokens=["Certificate", "Please refer", "2 of", "1 of"],
            )
            manufacturing = name_addr_block

        if not manufacturing:
            manufacturing = self._match(
                compact,
                [r"(Plot\s*No\.[^\.]{20,250}\.)"],
                "",
            )

        parts: list[str] = []
        if registered:
            parts.append("REGISTERED OFFICE:")
            parts.append(registered)
        if manufacturing:
            parts.append("FACTORY ADDRESS:")
            parts.append(manufacturing)
        name_and_address = "\n\n".join(parts).strip()

        responsibility = self._match(
            compact,
            [r"performs all\s+([^\.]+)\s+for this product", r"Responsibility\s+([\s\S]*?)\s+Name and Address"],
            "Manufacturing, testing",
        )

        return P31ManufacturerInfo(
            section_heading="2.3.P.3.1 Manufacturer(s)",
            name_and_address=name_and_address,
            responsibility=responsibility,
        )

    def extract(self, pdf_path: Path) -> ApiInfo:
        text = self._read_text(pdf_path)
        api_name = self._extract_api_name(text)
        manufacturer_text = self._extract_manufacturer_text(text, api_name)
        return ApiInfo(api_name=api_name, manufacturer_text=manufacturer_text)

    @staticmethod
    def _read_text(pdf_path: Path) -> str:
        doc = fitz.open(pdf_path)
        try:
            pages: list[str] = []
            for i in range(doc.page_count):
                pages.append(ApiInfoExtractor._read_page_text(doc.load_page(i)))
            return "\n".join(pages)
        finally:
            doc.close()

    @staticmethod
    def _body_clip_rect(page: fitz.Page) -> fitz.Rect:
        rect = page.rect
        h = float(rect.height)
        top = h * BODY_CLIP_TOP_RATIO
        bottom = h * BODY_CLIP_BOTTOM_RATIO
        return fitz.Rect(rect.x0, rect.y0 + top, rect.x1, rect.y0 + bottom)

    @staticmethod
    def _read_page_text(page: fitz.Page) -> str:
        return page.get_text("text", sort=True, clip=ApiInfoExtractor._body_clip_rect(page))

    @staticmethod
    def _find_page_index(pages: list[str], needle: str) -> int | None:
        for index, page_text in enumerate(pages):
            if needle.lower() in page_text.lower():
                return index
        return None

    @staticmethod
    def _find_summary_page_index(pages: list[str]) -> int | None:
        best_index: int | None = None
        best_score = 0

        for index, page_text in enumerate(pages):
            lower = page_text.lower()
            score = 0

            if "quality information summary" in lower:
                score += 2
            if re.search(r"\b1\.4\.2\b", lower):
                score += 2
            if "summary of product information" in lower:
                score += 3
            if "administrative summary" in lower:
                score += 1

            if score > best_score:
                best_score = score
                best_index = index

        # Minimum confidence avoids selecting a TOC mention-only page.
        if best_score >= 5:
            return best_index
        return None

    @staticmethod
    def _clean(value: str) -> str:
        return re.sub(r"\s+", " ", value).strip(" .:\n\t")

    def _match(self, text: str, patterns: list[str], default: str) -> str:
        for pattern in patterns:
            m = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
            if m:
                return self._clean(m.group(1))
        return default

    def _rows_to_label_value_map(self, rows: list[list[str | None]]) -> dict[str, list[str]]:
        values_by_label: dict[str, list[str]] = {}
        for row in rows:
            raw_cells = [str(cell or "") for cell in row]
            if not any(raw_cells):
                continue

            split_index = max(1, len(raw_cells) // 2)
            left_cells = [self._clean(cell) for cell in raw_cells[:split_index] if self._clean(cell)]
            right_cells = [self._clean_multiline(cell) for cell in raw_cells[split_index:] if self._clean_multiline(cell)]

            if not left_cells or not right_cells:
                continue

            raw_label = self._clean(" ".join(left_cells))
            label_key = self._normalize_label(raw_label)
            if not label_key:
                continue

            if label_key not in values_by_label:
                values_by_label[label_key] = right_cells
        return values_by_label

    def _repair_merged_address_rows(self, values_by_label: dict[str, list[str]]) -> None:
        building_key = self._normalize_label("Building/PO Box number")
        road_key = self._normalize_label("Road/Street")
        plant_key = self._normalize_label("Plant/Zone")
        village_key = self._normalize_label("Village/suburb")

        building_values = values_by_label.get(building_key)
        if not building_values:
            return

        first_building_value = building_values[0].strip()
        if not first_building_value:
            return

        road_present = bool(values_by_label.get(road_key) and any(v.strip() for v in values_by_label[road_key]))
        plant_present = bool(values_by_label.get(plant_key) and any(v.strip() for v in values_by_label[plant_key]))
        if road_present and plant_present:
            return

        lines = [self._clean(line) for line in first_building_value.splitlines() if self._clean(line)]
        if len(lines) >= 3:
            values_by_label[building_key] = [lines[0]] + building_values[1:]
            if not road_present:
                values_by_label[road_key] = [lines[1]]
            if not plant_present:
                values_by_label[plant_key] = [lines[2]]
            if len(lines) >= 4:
                village_present = bool(values_by_label.get(village_key) and any(v.strip() for v in values_by_label[village_key]))
                if not village_present:
                    values_by_label[village_key] = [lines[3]]
            return

        if len(lines) != 1:
            return

        comma_parts = [part.strip() for part in first_building_value.split(",") if part.strip()]
        if len(comma_parts) < 3:
            return

        building_main = ", ".join(comma_parts[:-2]).strip()
        road_main = f"{comma_parts[-2]},"
        plant_main = comma_parts[-1]
        if building_main:
            values_by_label[building_key] = [building_main] + building_values[1:]
        if not road_present:
            values_by_label[road_key] = [road_main]
        if not plant_present:
            values_by_label[plant_key] = [plant_main]

    def _repair_town_province_rows(self, values_by_label: dict[str, list[str]]) -> None:
        town_key = self._normalize_label("Town/City")
        province_key = self._normalize_label("Province/State")

        town_values = values_by_label.get(town_key)
        if not town_values:
            return

        province_present = bool(values_by_label.get(province_key) and any(v.strip() for v in values_by_label[province_key]))
        if province_present:
            return

        town_first = town_values[0].strip()
        if not town_first:
            return

        lines = [self._clean(line) for line in town_first.splitlines() if self._clean(line)]
        if len(lines) >= 2:
            values_by_label[town_key] = [lines[0]] + town_values[1:]
            values_by_label[province_key] = [lines[1]]

    def _extract_related_values_from_table(self, rows: list[list[str | None]]) -> list[str]:
        for row in reversed(rows):
            cells = [self._clean(str(cell or "")) for cell in row]
            cells = [cell for cell in cells if cell]
            if len(cells) < 2:
                continue
            if not any("not applicable" in cell.lower() for cell in cells):
                continue

            cleaned = [re.sub(r"\s+", " ", cell).strip() for cell in cells]
            cleaned = [cell for cell in cleaned if cell]
            if len(cleaned) >= 4:
                return cleaned[:4]
        return []

    @staticmethod
    def _normalize_label(label: str) -> str:
        normalized = label.lower()
        normalized = normalized.replace("\n", " ")
        normalized = re.sub(r"(?<=\D)\d+\b", "", normalized)
        normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        return normalized

    @staticmethod
    def _clean_multiline(value: str) -> str:
        normalized_lines: list[str] = []
        for raw_line in value.splitlines():
            line = re.sub(r"[ \t]+", " ", raw_line).strip(" .:\t")
            if line:
                normalized_lines.append(line)
        return "\n".join(normalized_lines)

    def _block_between(self, lines: list[str], start_tokens: list[str], end_tokens: list[str]) -> str:
        start = None
        for i, line in enumerate(lines):
            low = line.lower()
            if any(token.lower() in low for token in start_tokens):
                start = i + 1
                break
        if start is None:
            return ""

        end = len(lines)
        for i in range(start, len(lines)):
            low = lines[i].lower()
            if any(token.lower() in low for token in end_tokens):
                end = i
                break

        block_lines: list[str] = []
        for ln in lines[start:end]:
            cleaned = self._clean(ln)
            if not cleaned:
                continue
            if re.fullmatch(r"\d+", cleaned):
                continue
            if re.search(r"\b\d+\s+of\s+\d+\b", cleaned.lower()):
                continue
            block_lines.append(cleaned)
        return "\n".join(block_lines)

    def _first_address_block(self, lines: list[str]) -> str:
        for i, line in enumerate(lines):
            low = line.lower()
            if "manufacturing facility-1" in low or "address of manufacturer" in low:
                address_lines: list[str] = []
                for nxt in lines[i + 1 : i + 9]:
                    nxt_clean = self._clean(nxt)
                    if not nxt_clean:
                        break
                    if "manufacturing facility-2" in nxt_clean.lower() or "certificate" in nxt_clean.lower():
                        break
                    address_lines.append(nxt_clean)
                if address_lines:
                    return " ".join(address_lines)
        return ""

    def _company_from_manufactured_by(self, lines: list[str]) -> str:
        for i, line in enumerate(lines):
            low = line.lower()
            if "manufactured" not in low or " by " not in low:
                continue
            tail = re.split(r"\bby\b", line, flags=re.IGNORECASE, maxsplit=1)[-1].strip()
            candidate = tail
            if candidate.endswith("."):
                candidate = candidate[:-1].strip()
            # If split line ends too early (e.g., 'Zhejiang' or 'M/s.'), append next line.
            if (len(candidate.split()) <= 2 or candidate.lower() in {"m/s", "m/s."}) and i + 1 < len(lines):
                candidate = f"{candidate} {self._clean(lines[i + 1])}".strip()
            candidate = self._clean(candidate)
            if candidate:
                return candidate
        return ""

    def _extract_api_name(self, text: str) -> str:
        patterns = [
            r"active\s+drug\s+([A-Za-z0-9\-\s\(\)]+?)\s+is manufactured",
            r"active\s+pharmaceutical\s+ingredient\s*[-:]?\s*([A-Za-z0-9\-\s\(\)]+?)\s+is manufactured",
            r"drug\s+substance\s*[:\-]?\s*([A-Za-z0-9\-\s\(\)]+)",
        ]
        for pattern in patterns:
            m = re.search(pattern, text, flags=re.IGNORECASE)
            if m:
                return self._clean(m.group(1))
        return ""

    def _extract_manufacturer_text(self, text: str, api_name: str) -> str:
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        compact = re.sub(r"\s+", " ", text)

        main_company = self._match(
            compact,
            [r"is manufactured[^\n\.]*?by\s+([^\.\n]+)"],
            "",
        )
        inter_sentence = self._match(
            compact,
            [r"([A-Za-z0-9\-\s\(\),]+intermediate[^\n\.]*?is manufactured[^\n\.]*?by\s+[^\.\n]+)"],
            "",
        )
        inter_company = self._match(
            compact,
            [r"intermediate[^\n\.]*?is manufactured[^\n\.]*?by\s+([^\.\n]+)"],
            "",
        )

        inter_address = ""
        if inter_company:
            for i, ln in enumerate(lines):
                if inter_company.lower() in self._clean(ln).lower():
                    chunks: list[str] = []
                    for nxt in lines[i + 1 : i + 5]:
                        low = nxt.lower()
                        if "certificate" in low or "page" in low:
                            break
                        chunks.append(self._clean(nxt))
                    inter_address = " ".join([c for c in chunks if c]).strip()
                    break

        parts: list[str] = []
        if main_company:
            parts.append(main_company)
        if inter_sentence:
            if parts:
                parts.append("")
            parts.append(inter_sentence + ".")
        if inter_company:
            if parts:
                parts.append("")
            parts.append(inter_company)
        if inter_address:
            parts.append(inter_address)

        if not parts and api_name:
            return ""

        return "\n".join(parts).strip()
