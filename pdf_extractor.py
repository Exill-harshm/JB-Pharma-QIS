"""
Module: pdf_extractor
Responsibility: Extracts content from CTD Module 3 PDFs.
"""
import os
import fitz  # PyMuPDF
import warnings
from collections import Counter
from typing import List, Optional, Dict, Set
from logger_setup import get_logger

warnings.filterwarnings("ignore")

MIN_CONTENT_CHARS = 100
BODY_CLIP_TOP_RATIO = 0.10
BODY_CLIP_BOTTOM_RATIO = 0.92
MIN_BODY_WORDS = 3


class ExtractedContent:
    """Structured result from a single PDF extraction."""
    def __init__(self):
        self.docx_path:       str        = ""
        self.noise_blocklist: Set[str]   = set()


def _build_noise_blocklist(pdf_path: str, logger) -> Set[str]:
    """
    Auto-detects header/footer text by finding text that repeats
    across multiple pages in the top/bottom margins of the PDF.

    Algorithm:
    - For each page, collect normalised text from the top 12% and bottom 10%.
    - Any text appearing on >= min(3, total_pages) pages is treated as noise.

    Returns a set of normalised lowercase strings to suppress during cleaning.
    """
    try:
        doc         = fitz.open(pdf_path)
        total_pages = len(doc)

        if total_pages <= 1:
            doc.close()
            return set()

        # Collect per-page sets so a string appearing in both top AND bottom
        # of the same page is only counted once per page.
        page_texts: List[Set[str]] = []
        for page in doc:
            h        = page.rect.height
            w        = page.rect.width
            page_set: Set[str] = set()
            clips = [
                fitz.Rect(0, 0,        w, h * 0.12),  # top 12 %
                fitz.Rect(0, h * 0.90, w, h),          # bottom 10 %
            ]
            for clip in clips:
                for block in page.get_text("blocks", clip=clip):
                    text = block[4].strip()
                    if not text:
                        continue
                    norm = " ".join(text.lower().split())
                    if len(norm) >= 3:
                        page_set.add(norm)
            page_texts.append(page_set)

        doc.close()

        # Count how many pages each text appears on
        freq: Counter = Counter()
        for page_set in page_texts:
            for text in page_set:
                freq[text] += 1

        # Text on >= threshold pages is noise (header/footer)
        threshold = min(3, total_pages)
        blocklist = {t for t, c in freq.items() if c >= threshold}

        if blocklist:
            logger.info(
                f"{os.path.basename(pdf_path)}: "
                f"auto-detected {len(blocklist)} header/footer noise strings."
            )

        return blocklist

    except Exception as e:
        logger.warning(
            f"Could not build noise blocklist for "
            f"{os.path.basename(pdf_path)}: {e}"
        )
        return set()


def _detect_with_layout(pdf_path: str, logger) -> Optional[List[int]]:
    """
    AI-based header/footer removal using pymupdf-layout + pymupdf4llm.

    Returns 0-based content page indices. Returns None on any failure.
    """
    try:
        import pymupdf.layout
        import pymupdf4llm

        chunks = pymupdf4llm.to_markdown(
            pdf_path,
            page_chunks=True,
            header=False,
            footer=False,
            show_progress=False,
        )

        if not isinstance(chunks, list):
            logger.warning(
                f"{os.path.basename(pdf_path)}: page_chunks ignored "
                f"(got string). Upgrade pymupdf4llm. Using fallback."
            )
            return None

        content_pages = []
        for chunk in chunks:
            page_idx  = chunk.get("metadata", {}).get("page", 0)
            text      = chunk.get("text", "")
            real_text = " ".join(text.split())
            if len(real_text) >= MIN_CONTENT_CHARS:
                content_pages.append(page_idx)

        return content_pages if content_pages else None

    except ImportError as e:
        logger.warning(
            f"Layout import failed ({e}). "
            f"Install: pip install pymupdf-layout. Using fallback."
        )
        return None
    except Exception as e:
        logger.error(
            f"Layout detection failed for "
            f"{os.path.basename(pdf_path)}: {e}"
        )
        return None


def _detect_with_fallback(pdf_path: str, logger) -> Optional[List[int]]:
    """
    Fallback detection using percentage-based margin clipping.
    """
    try:
        doc           = fitz.open(pdf_path)
        content_pages = []
        for i, page in enumerate(doc):
            h    = page.rect.height
            w    = page.rect.width
            clip = fitz.Rect(0, h * 0.22, w, h * 0.90)
            text = page.get_text(clip=clip).strip()
            if len(" ".join(text.split())) >= MIN_CONTENT_CHARS:
                content_pages.append(i)
        doc.close()
        return content_pages if content_pages else None
    except Exception as e:
        logger.error(f"Fallback detection failed for {pdf_path}: {e}")
        return None


def _detect_content_pages(pdf_path: str, logger) -> Optional[List[int]]:
    """
    Detect content pages using layout mode first, then fallback clipping.
    """
    result = _detect_with_layout(pdf_path, logger)
    if result is not None:
        # Prevent bug where pymupdf4llm chunks all return page 0, truncating the PDF.
        if len(result) > 1 and all(p == 0 for p in result):
            logger.warning(
                f"{os.path.basename(pdf_path)}: AI chunks lost page indices. Using fallback."
            )
        else:
            # Reconstruct unique sorted page list to fix duplicate chunk indices
            return sorted(list(set(result)))
            
    logger.warning(
        f"{os.path.basename(pdf_path)}: Using fallback clip-based page detection."
    )
    return _detect_with_fallback(pdf_path, logger)


def _normalize_text(value: str) -> str:
    return " ".join((value or "").lower().split())


def _body_clip_rect(page, blocklist: Set[str]) -> fitz.Rect:
    """
    Computes a robust page body rectangle by excluding likely header/footer bands.
    Uses ratio defaults plus repeated-margin text hints from the blocklist.
    """
    page_rect = page.rect
    h = float(page_rect.height)

    default_top = h * BODY_CLIP_TOP_RATIO
    default_bottom = h * BODY_CLIP_BOTTOM_RATIO
    top = default_top
    bottom = default_bottom

    try:
        header_bottoms: List[float] = []
        footer_tops: List[float] = []
        for block in page.get_text("blocks"):
            if len(block) < 5:
                continue
            x0, y0, x1, y1, text = block[:5]
            norm = _normalize_text(text)
            if not norm or norm not in blocklist:
                continue

            # Repeated strings near top are header candidates.
            if y1 <= h * 0.28:
                header_bottoms.append(float(y1))

            # Repeated strings near bottom are footer candidates.
            if y0 >= h * 0.72:
                footer_tops.append(float(y0))

        if header_bottoms:
            top = max(top, max(header_bottoms) + 2.0)
        if footer_tops:
            bottom = min(bottom, min(footer_tops) - 2.0)
    except Exception:
        pass

    # Guardrails: avoid over-cropping pages with unusual layouts.
    top = max(0.0, min(top, h * 0.45))
    bottom = min(h, max(bottom, h * 0.55))
    if bottom - top < h * 0.35:
        top = default_top
        bottom = default_bottom

    return fitz.Rect(page_rect.x0, page_rect.y0 + top, page_rect.x1, page_rect.y0 + bottom)


def _page_has_body_content(page, body_rect: fitz.Rect) -> bool:
    """Returns True if clipped body contains enough textual content to keep the page."""
    try:
        words = page.get_text("words", clip=body_rect)
        if len(words) >= MIN_BODY_WORDS:
            return True
        text = page.get_text("text", clip=body_rect)
        return len(" ".join(text.split())) >= MIN_CONTENT_CHARS
    except Exception:
        return False


def _build_body_clipped_pdf(
    pdf_path: str,
    page_indices: List[int],
    blocklist: Set[str],
    log_folder: str,
    base_name: str,
    logger,
) -> tuple[Optional[str], List[int]]:
    """
    Creates a temporary PDF composed of clipped body regions only.
    Pages that become header/footer-only after clipping are dropped.
    """
    src_doc = fitz.open(pdf_path)
    dst_doc = fitz.open()
    kept_pages: List[int] = []
    temp_pdf_path: Optional[str] = None

    try:
        for page_index in page_indices:
            if page_index < 0 or page_index >= src_doc.page_count:
                continue

            page = src_doc[page_index]
            body_rect = _body_clip_rect(page, blocklist)

            if not _page_has_body_content(page, body_rect):
                logger.info(
                    f"{base_name}: dropping page {page_index + 1} "
                    f"(header/footer-only after clipping)."
                )
                continue

            out_page = dst_doc.new_page(
                width=body_rect.width,
                height=body_rect.height,
            )
            out_page.show_pdf_page(out_page.rect, src_doc, page_index, clip=body_rect)
            kept_pages.append(page_index)

        if not kept_pages:
            return None, []

        temp_pdf_path = os.path.join(log_folder, f"temp_bodyclip_{base_name}")
        dst_doc.save(temp_pdf_path, garbage=4, deflate=True)
        logger.info(
            f"{base_name}: body-clipped PDF created with {len(kept_pages)} page(s)."
        )
        return temp_pdf_path, kept_pages
    finally:
        dst_doc.close()
        src_doc.close()


def extract_pdf_content(
    pdf_path:            str,
    log_folder:          str,
    section_num:         str             = "",
    section_page_limits: Dict[str, int] = None,
    section_start_pages: Dict[str, int] = None,
) -> ExtractedContent:
    """
    Main extraction entry point.
    """
    logger    = get_logger(log_folder)
    content   = ExtractedContent()
    base_name = os.path.basename(pdf_path)

    if section_page_limits is None:
        section_page_limits = {}
    if section_start_pages is None:
        section_start_pages = {}

    # Build auto-detected noise blocklist FIRST (used later in docx_builder)
    content.noise_blocklist = _build_noise_blocklist(pdf_path, logger)

    content_pages = _detect_content_pages(pdf_path, logger)
    if content_pages:
        logger.info(
            f"{base_name}: {len(content_pages)} content pages "
            f"(0-based: {content_pages})"
        )
    else:
        logger.warning(
            f"{base_name}: no content pages detected — using all pages."
        )

    # Determine page indices requested by section config.
    try:
        doc_probe = fitz.open(pdf_path)
        total_pages = doc_probe.page_count
        doc_probe.close()
    except Exception:
        total_pages = 0

    start_page = 0
    if section_num and section_num in section_start_pages:
        start_page = max(0, int(section_start_pages[section_num]))
        logger.info(
            f"{base_name}: skipping first {start_page} pages "
            f"(cover/TOC). Starting at page {start_page + 1}."
        )

    end_page = total_pages
    if section_num and section_num in section_page_limits:
        limit = max(1, int(section_page_limits[section_num]))
        end_page = min(total_pages, start_page + limit)
        logger.info(
            f"{base_name}: converting pages {start_page} to {end_page} only."
        )

    requested_pages = list(range(start_page, end_page)) if total_pages else []

    if content_pages:
        content_page_set = set(content_pages)
        filtered = [p for p in requested_pages if p in content_page_set]
        if filtered:
            requested_pages = filtered

    if not requested_pages and total_pages:
        requested_pages = list(range(total_pages))

    logger.info(
        f"{base_name}: requested page indices after filtering: {requested_pages}"
    )

    logger.info(f"Converting {base_name} via pdf2docx")
    try:
        from pdf2docx import Converter
        temp_docx_path = os.path.join(
            log_folder, f"temp_layout_{base_name}.docx"
        )

        body_pdf_path, kept_pages = _build_body_clipped_pdf(
            pdf_path=pdf_path,
            page_indices=requested_pages,
            blocklist=content.noise_blocklist,
            log_folder=log_folder,
            base_name=base_name,
            logger=logger,
        )

        if not body_pdf_path or not kept_pages:
            logger.warning(
                f"{base_name}: no pages left after header/footer clipping."
            )
            return content

        cv = Converter(body_pdf_path)
        try:
            cv.convert(
                temp_docx_path,
                start=0,
                end=None,
                multi_processing=False,
            )
        finally:
            cv.close()

        try:
            os.remove(body_pdf_path)
        except Exception:
            pass

        if os.path.exists(temp_docx_path):
            content.docx_path = temp_docx_path
            logger.info(f"pdf2docx converted {base_name} OK.")
        else:
            logger.error(f"pdf2docx no output for {base_name}.")
    except Exception as e:
        logger.error(
            f"pdf2docx failed for {base_name}: {e}", exc_info=True
        )

    logger.info(f"Image extraction disabled for {base_name}.")

    return content