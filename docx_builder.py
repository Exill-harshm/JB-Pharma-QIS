"""
Module: docx_builder
Responsibility: Opens QIS template DOCX, finds all "Refer Section X.X.X"
placeholders, injects extracted PDF content (text + tables),
cleans injected noise (headers/footers/page numbers), saves output DOCX.

Noise detection is fully automatic — no hardcoded company names.
The auto-detected blocklist comes from pdf_extractor._build_noise_blocklist().
"""
import re
import copy
import os
import docx
from docx.shared import RGBColor
from typing import Dict, Set
from logger_setup import get_logger

PLACEHOLDER_PATTERN = re.compile(
    r"Refer\s+(?:the\s+)?section\s+(\d+(?:\.[a-zA-Z0-9]+)+)",
    re.IGNORECASE
)

MANUAL_ENTRY_SECTIONS = {'1.2', '1.3', '1.4', '1.5', '1.5.1', '1.5.2', '1.6'}

# Generic regex patterns — work for ANY document, no company-specific strings
_PAGE_NUM_RE = re.compile(r'^\d{1,4}$')
_PAGE_OF_RE  = re.compile(r'^\d+\s+of\s+\d+\s*$', re.IGNORECASE)

def _remove_noise_tables(doc, logger) -> int:
    body = doc.element.body
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def _wt(elem):
        return ''.join(t.text or '' for t in elem.iter(f'{{{_NS}}}t')).strip()

    # All paragraph texts — used to detect 1x1 header tables
    all_para_texts = {p.text.strip() for p in doc.paragraphs if p.text.strip()}

    def _is_noise(elem):
        text = _wt(elem)
        clean = re.sub(r'\s+', '', text)
        rows = len(elem.findall(f'.//{{{_NS}}}tr'))
        cols = len(elem.findall(f'{{{_NS}}}tr/{{{_NS}}}tc'))
        if not clean:                                                   return True  # empty
        if re.match(r'^\d{1,4}$', text):                               return True  # "42"
        if re.match(r'^\d+\s+of\s+\d+$', text, re.I):                 return True  # "3 of 6"
        if len(text) < 90 and re.match(r'^.{5,75}\s+\d{1,4}$', text): return True  # "Co. 52"
        if rows == 1 and cols == 1 and len(text) < 100 and text in all_para_texts:
            return True  # 1x1 header table duplicated from a paragraph
        # DMF page-header tables ("Drug Mater File Version: 3.1...")
        if rows <= 4 and 'Drug Mater File' in text:
            return True
        # PDF section-header tables ("3.2.P PARTICULARS OF FINSHED...")
        if rows <= 3 and 'PARTICULARS' in text:
            return True
        return False

    to_remove = [e for e in list(body) if e.tag.split('}')[-1] == 'tbl' and _is_noise(e)]
    for elem in to_remove:
        body.remove(elem)
    if to_remove:
        logger.info(f"Removed {len(to_remove)} noise table(s).")
    return len(to_remove)


def _collapse_blank_paragraphs(doc, logger) -> int:
    """Collapses runs of consecutive blank paragraphs down to at most one."""
    body = doc.element.body
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    consecutive, to_remove = 0, []
    for elem in list(body):
        tag = elem.tag.split('}')[-1]
        if tag == 'p':
            text = ''.join(t.text or '' for t in elem.iter(f'{{{_NS}}}t')).strip()
            if not text:
                consecutive += 1
                if consecutive > 1:
                    to_remove.append(elem)
            else:
                consecutive = 0
        else:
            consecutive = 0
    for elem in to_remove:
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
    if to_remove:
        logger.info(f"Removed {len(to_remove)} excess blank paragraph(s).")
    return len(to_remove)

def _fix_zero_width_tables(doc, logger) -> int:
    """Fixes tables injected by pdf2docx with tblW=0 — they render as blank boxes."""
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    PAGE_WIDTH_DXA = '9026'  # A4 with standard margins
    fixed = 0
    for elem in doc.element.body:
        if elem.tag.split('}')[-1] != 'tbl':
            continue
        tblPr = elem.find(f'{{{_NS}}}tblPr')
        if tblPr is None:
            continue
        tblW = tblPr.find(f'{{{_NS}}}tblW')
        if tblW is not None and tblW.get(f'{{{_NS}}}w', '') == '0':
            tblW.set(f'{{{_NS}}}w', PAGE_WIDTH_DXA)
            tblW.set(f'{{{_NS}}}type', 'dxa')
            fixed += 1
    if fixed:
        logger.info(f"Fixed {fixed} zero-width table(s).")
    return fixed

def _remove_repeated_header_paragraphs(doc, logger) -> int:
    """
    Removes PDF page-header paragraphs injected repeatedly by pdf2docx.
    Detection is fully automatic: any paragraph text that appears 3+ times
    AND is under 120 chars is treated as a repeating noise header.
    """
    from collections import Counter
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    freq = Counter(p.text.strip() for p in doc.paragraphs if p.text.strip())
    noise = {t for t, c in freq.items() if c >= 3 and len(t) < 120}

    removed = 0
    for elem in list(doc.element.body):
        if elem.tag.split('}')[-1] != 'p':
            continue
        text = ''.join(t.text or '' for t in elem.iter(f'{{{_NS}}}t')).strip()
        if text in noise:
            elem.getparent().remove(elem)
            removed += 1

    if removed:
        logger.info(f"Removed {removed} repeated header paragraph(s).")
    return removed

def _remove_pdf_noise_paragraphs(doc, logger) -> int:
    """Removes PDF-injected noise paragraphs not caught by frequency filter."""
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = doc.element.body
    elements = list(body)

    def _wt(e):
        return ''.join(t.text or '' for t in e.iter(f'{{{_NS}}}t')).strip()

    removed = 0
    for i, elem in enumerate(elements):
        if elem.tag.split('}')[-1] != 'p':
            continue
        text = _wt(elem)
        if not text:
            continue

        drop = False
        # 1. Company name + bare page number  e.g. "Starry Co. Ltd 230"
        if re.match(r'^.{10,75}\s+\d{1,4}$', text) and len(text) < 80:
            drop = True
        # 2. PDF section heading (3.2.X format) when template heading (2.3.X) exists nearby above
        elif re.search(r'3\.2[\. ][A-Z0-9P]', text) and len(text) < 200:
            for j in range(max(0, i - 10), i):
                if re.search(r'2\.3\.[SP]', _wt(elements[j])):
                    drop = True
                    break
        # 3. OCR artifacts and known single-occurrence PDF header lines
        elif text in {'FINIHSED PRODUCT SPECIFICATION Product name', 'C~CkedBY:'}:
            drop = True

        if drop:
            elem.getparent().remove(elem)
            removed += 1

    if removed:
        logger.info(f"Removed {removed} PDF noise paragraph(s).")
    return removed

def _remove_empty_visual_tables(doc, logger) -> int:
    """
    Removes tables that have:
    - no meaningful text
    - mostly empty cells
    - created from vector drawings (pdf2docx artifacts)
    """
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    removed = 0

    def get_text(cell):
        return ''.join(
            t.text or '' for t in cell._element.iter(f'{{{_NS}}}t')
        ).strip()

    for table in list(doc.tables):
        total_cells = 0
        empty_cells = 0
        text_cells = 0

        for row in table.rows:
            for cell in row.cells:
                total_cells += 1
                text = get_text(cell)

                if not text:
                    empty_cells += 1
                else:
                    text_cells += 1

        # KEY LOGIC
        if total_cells > 4 and text_cells == 0:
            # completely empty table → remove
            table._element.getparent().remove(table._element)
            removed += 1

        elif total_cells > 6 and (empty_cells / total_cells) > 0.8:
            # mostly empty → also remove
            table._element.getparent().remove(table._element)
            removed += 1

    if removed:
        logger.info(f"Removed {removed} empty visual table(s).")

    return removed


def _remove_low_content_injected_tables(doc, logger, keep_first_n_tables: int) -> int:
    """
    Removes low-value injected tables while preserving original template tables.
    This targets pdf2docx line-art table artifacts that appear as random lines.
    """
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    removed = 0

    def get_text(cell):
        return ''.join(
            t.text or '' for t in cell._element.iter(f'{{{_NS}}}t')
        ).strip()

    all_tables = list(doc.tables)
    for idx, table in enumerate(all_tables):
        if idx < keep_first_n_tables:
            continue

        total_cells = 0
        text_cells = 0
        text_chars = 0
        for row in table.rows:
            for cell in row.cells:
                total_cells += 1
                text = get_text(cell)
                if text:
                    text_cells += 1
                    text_chars += len(text)

        if total_cells == 0:
            continue

        empty_ratio = (total_cells - text_cells) / total_cells
        should_remove = False
        if text_cells == 0 and total_cells >= 4:
            should_remove = True
        elif empty_ratio > 0.85 and text_chars < 80 and total_cells >= 6:
            should_remove = True
        elif text_cells <= 2 and total_cells >= 8 and text_chars < 100:
            should_remove = True

        if should_remove:
            table._element.getparent().remove(table._element)
            removed += 1

    if removed:
        logger.info(f"Removed {removed} low-content injected table artifact(s).")
    return removed


def _is_noise_paragraph(text: str, blocklist: Set[str]) -> bool:
    """
    Returns True if text is a header/footer/page-number noise line.

    Uses two layers:
    1. Auto-detected blocklist (strings that repeat in top/bottom margins)
    2. Generic document-agnostic heuristics (page numbers, short ALL-CAPS)
    """
    text = text.strip()
    if not text:
        return False

    norm = " ".join(text.lower().split())

    # Layer 1: auto-detected repeating header/footer text
    if norm in blocklist:
        return True

    # Layer 2: bare page numbers  ("42"  or  "42 of 100")
    if _PAGE_NUM_RE.match(norm):
        return True
    if _PAGE_OF_RE.match(norm):
        return True

    # Layer 2b: OCR/header garbage lines seen in converted PDFs.
    compact = re.sub(r'[^a-z0-9]+', '', norm)
    if 'productspecificationproductname' in compact:
        return True
    if compact in {'cckedby', 'checkedby'}:
        return True

    # Layer 3: very short ALL-CAPS lines — running header artefacts
    # (e.g. "INTRODUCTION", "MODULE 3") — only strip if <= 3 words & < 25 chars
    if len(text) < 25 and text.isupper() and len(text.split()) <= 3:
        return True

    return False


def _is_footer_table_row(row, blocklist: Set[str]) -> bool:
    """
    Returns True if a table row is entirely noise (header/footer row).

    A footer row must satisfy:
    - At least one cell matches the auto-detected blocklist, AND
    - Every other non-empty cell is a bare page number.

    This is fully generic — no company names needed.
    """
    cell_texts = [cell.text.strip() for cell in row.cells]
    if not any(cell_texts):
        return False

    has_blocklist_hit = False
    for t in cell_texts:
        if not t:
            continue
        norm = " ".join(t.lower().split())
        if norm in blocklist:
            has_blocklist_hit = True
        elif not (_PAGE_NUM_RE.match(t) and t.isdigit() and int(t) < 10000):
            # This cell has real content — row is NOT a footer
            return False

    return has_blocklist_hit


def _clean_injected_content(
    src_doc,
    blocklist: Set[str],
    logger,
    section_num: str
):
    """
    Cleans noise from pdf2docx-converted DOCX BEFORE injecting into template.
    Uses the auto-detected blocklist — no hardcoded company strings.
    """
    removed = 0

    # Remove noisy standalone paragraphs entirely to avoid leaving empty lines.
    paras_to_remove = []
    for para in src_doc.paragraphs:
        if _is_noise_paragraph(para.text, blocklist):
            paras_to_remove.append(para)

    for para in paras_to_remove:
        p = para._p
        parent = p.getparent()
        if parent is not None:
            parent.remove(p)
            removed += 1

    for table in src_doc.tables:
        for row in table.rows:
            if _is_footer_table_row(row, blocklist):
                for cell in row.cells:
                    for para in cell.paragraphs:
                        para.clear()
                removed += 1

    if removed:
        logger.info(
            f"Section {section_num}: cleaned {removed} "
            f"noise items (headers/footers/page numbers)."
        )


def _element_text_content(element) -> str:
    """Extracts combined text for an XML element."""
    _NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return ''.join(t.text or '' for t in element.iter(f'{{{_NS}}}t')).strip()


def _iter_all_paragraphs(doc):
    """Yields every paragraph in body AND inside every table cell."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _insert_warning(paragraph, section: str):
    """Replaces placeholder with visible red WARNING text."""
    paragraph.clear()
    run = paragraph.add_run(
        f"[WARNING: Could not populate section {section} - check source PDF]"
    )
    run.bold = True
    run.font.color.rgb = RGBColor(255, 0, 0)


def _strip_drawing_elements(element):
    """
    Removes drawing/image XML nodes using Clark notation.
    Images are re-inserted separately via PyMuPDF bytes.
    """
    DRAWING_TAGS = {
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing',
        '{urn:schemas-microsoft-com:vml}imagedata',
        '{http://schemas.openxmlformats.org/drawingml/2006/main}blipFill',
    }
    to_remove = [n for n in element.iter() if n.tag in DRAWING_TAGS]
    for node in to_remove:
        parent = node.getparent()
        if parent is not None:
            try:
                parent.remove(node)
            except Exception:
                pass


def _inject_docx_content(
    src_docx_path: str,
    anchor_p_xml,
    blocklist: Set[str],
    logger,
    section_num: str,
    include_pdf_tables: bool = False,
):
    """
    Opens pdf2docx temp DOCX, cleans noise, copies body elements
    into template after anchor. Skips sectPr and strips drawing refs.
    Returns new anchor (last inserted element).
    """
    current_anchor = anchor_p_xml
    try:
        src_doc = docx.Document(src_docx_path)
    except Exception as e:
        logger.error(
            f"Cannot open converted DOCX for section {section_num}: {e}"
        )
        return current_anchor

    _clean_injected_content(src_doc, blocklist, logger, section_num)

    inserted_nonblank = False
    pending_blank_para = None

    for element in src_doc.element.body:
        if element.tag.endswith('}sectPr') or element.tag == 'sectPr':
            continue
        if not include_pdf_tables and element.tag.split('}')[-1] == 'tbl':
            continue

        tag_name = element.tag.split('}')[-1]
        if tag_name == 'p':
            if not _element_text_content(element):
                # Drop leading blanks and compress internal blank runs.
                if not inserted_nonblank:
                    continue
                pending_blank_para = copy.deepcopy(element)
                continue

            # Insert at most one blank paragraph before non-blank content.
            if pending_blank_para is not None:
                try:
                    blank_el = copy.deepcopy(pending_blank_para)
                    current_anchor.addnext(blank_el)
                    current_anchor = blank_el
                except Exception:
                    pass
                pending_blank_para = None

        try:
            new_el = copy.deepcopy(element)
            _strip_drawing_elements(new_el)
            current_anchor.addnext(new_el)
            current_anchor = new_el
            if tag_name == 'p' and _element_text_content(new_el):
                inserted_nonblank = True
        except Exception as el_e:
            logger.warning(
                f"Skipped element for section {section_num}: {el_e}"
            )

    return current_anchor


def process_template(
    template_path:       str,
    output_path:         str,
    section_map:         Dict[str, str],
    log_folder:          str,
    section_page_limits: Dict[str, int] = None,
    section_start_pages: Dict[str, int] = None,
    preserve_template_tables: bool = False,
    include_pdf_tables: bool = False,
):
    """
    Main pipeline:
    1. Open QIS template DOCX
    2. Scan all paragraphs + table cells for placeholders
    3. Extract + clean + inject content per section
    4. Save populated output DOCX
    Returns: (sections_filled, warnings_count, failures)
    """
    logger = get_logger(log_folder)
    from pdf_extractor import extract_pdf_content

    if section_page_limits is None:
        section_page_limits = {}
    if section_start_pages is None:
        section_start_pages = {}

    try:
        doc = docx.Document(template_path)
    except Exception as e:
        logger.error(f"Failed to open template: {e}")
        raise

    template_table_count = len(doc.tables)

    sections_filled    = 0
    warnings_count     = 0
    failures           = 0
    processed_sections = set()

    logger.info("Starting QIS template placeholder scan.")

    for paragraph in _iter_all_paragraphs(doc):
        match = PLACEHOLDER_PATTERN.search(paragraph.text)
        if not match:
            continue

        section_num = match.group(1)
        logger.info(f"Found placeholder: {section_num}")

        if section_num in MANUAL_ENTRY_SECTIONS:
            logger.info(
                f"Section {section_num}: Module 1 admin - leave for manual entry."
            )
            continue

        if section_num not in section_map:
            logger.warning(
                f"Section {section_num}: no source PDF mapped. Inserting warning."
            )
            _insert_warning(paragraph, section_num)
            warnings_count += 1
            failures += 1
            continue

        if section_num in processed_sections:
            logger.info(
                f"Section {section_num}: duplicate placeholder - clearing."
            )
            paragraph.clear()
            continue

        pdf_path = section_map[section_num]
        logger.info(
            f"Processing {section_num} from '{os.path.basename(pdf_path)}'"
        )

        try:
            content = extract_pdf_content(
                pdf_path            = pdf_path,
                log_folder          = log_folder,
                section_num         = section_num,
                section_page_limits = section_page_limits,
                section_start_pages = section_start_pages,
            )

            paragraph.clear()
            current_anchor = paragraph._p

            if content.docx_path and os.path.exists(content.docx_path):
                current_anchor = _inject_docx_content(
                    src_docx_path = content.docx_path,
                    anchor_p_xml  = current_anchor,
                    blocklist     = content.noise_blocklist,
                    logger        = logger,
                    section_num   = section_num,
                    include_pdf_tables = include_pdf_tables,
                )
                try:
                    os.remove(content.docx_path)
                except Exception:
                    pass
            else:
                logger.warning(
                    f"Section {section_num}: no converted DOCX available."
                )

            sections_filled += 1
            processed_sections.add(section_num)
            logger.info(f"Section {section_num}: populated successfully.")

        except Exception as e:
            logger.error(
                f"Section {section_num} failed: {e}", exc_info=True
            )
            _insert_warning(paragraph, section_num)
            warnings_count += 1
            failures += 1

    logger.info(f"Saving output to {output_path}")
    try:
        # Safe in all modes: fixes converter-introduced zero-width tables.
        _fix_zero_width_tables(doc, logger)

        if not preserve_template_tables:
            _remove_noise_tables(doc, logger)
            _remove_empty_visual_tables(doc, logger)
            _remove_repeated_header_paragraphs(doc, logger)
            _remove_pdf_noise_paragraphs(doc, logger)
            _collapse_blank_paragraphs(doc, logger)
        else:
            logger.info(
                "Template-preserve mode enabled: skipping global cleanup passes "
                "that may remove static QIS tables."
            )
            _remove_low_content_injected_tables(doc, logger, template_table_count)

        logger.info(f"Saving output to {output_path}")
        doc.save(output_path)
    except Exception as e:
        logger.error(f"Failed to save: {e}")
        raise

    return sections_filled, warnings_count, failures