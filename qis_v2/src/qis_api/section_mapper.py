from __future__ import annotations

import re
from pathlib import Path


class SectionMapper:
    def __init__(self, dossier_root: Path):
        self._dossier_root = dossier_root

    def resolve_pdf(self, ctd_reference: str) -> Path | None:
        pdfs = list(self._dossier_root.rglob("*.pdf"))
        if not pdfs:
            return None

        ref = ctd_reference.strip()
        while ref:
            exact = self._find_exact(pdfs, ref)
            if exact is not None:
                return exact

            prefix = self._find_prefix(pdfs, ref)
            if prefix is not None:
                return prefix

            ref = self._parent_ref(ref)

        return None

    @staticmethod
    def _normalize_stem(stem: str) -> str:
        lowered = stem.lower().replace(" ", "")
        lowered = lowered.replace("-", "")
        lowered = lowered.replace("_", "")
        return lowered

    def _find_exact(self, pdfs: list[Path], ref: str) -> Path | None:
        target = self._normalize_stem(ref)
        for pdf in pdfs:
            if self._normalize_stem(pdf.stem) == target:
                return pdf
        return None

    def _find_prefix(self, pdfs: list[Path], ref: str) -> Path | None:
        target = self._normalize_stem(ref)
        pattern = re.compile(rf"(^|[^0-9]){re.escape(target)}([^0-9]|$)")
        for pdf in pdfs:
            stem = self._normalize_stem(pdf.stem)
            if stem.startswith(target) or pattern.search(stem):
                return pdf
        return None

    @staticmethod
    def _parent_ref(ref: str) -> str:
        if "." not in ref:
            return ""
        return ref.rsplit(".", 1)[0]
