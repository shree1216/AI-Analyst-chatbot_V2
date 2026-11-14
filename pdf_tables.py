"""Turn tables found in a PDF into a workbook so PDFs can be analysed like spreadsheets."""

from __future__ import annotations

from io import BytesIO
import re

import pandas as pd

from data_explorer import _clean_headers, _numeric_column
from excel_support import MAX_EXCEL_BYTES, MAX_EXCEL_CELLS, MAX_UPLOAD_MB, generate_excel_download

MAX_PDF_PAGES = 200
MAX_PDF_TABLES = 25


def _cell(value) -> str:
    return re.sub(r"\s+", " ", str(value)).strip() if value is not None else ""


def _frame(rows: list[list]) -> pd.DataFrame | None:
    rows = [[_cell(value) for value in row] for row in rows if row and any(_cell(value) for value in row)]
    if len(rows) < 2:
        return None
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    frame = pd.DataFrame(rows[1:], columns=_clean_headers(rows[0]), dtype=str)
    frame = frame.loc[:, [column for column in frame.columns if frame[column].str.strip().ne("").any()]]
    if frame.empty or frame.shape[1] < 2:
        return None
    for column in frame.columns:
        cleaned = frame[column].str.replace(r"(?<=\d),(?=\d{3}\b)", "", regex=True)
        converted = _numeric_column(cleaned)
        frame[column] = converted if converted is not None else frame[column].where(frame[column] != "", None)
    return frame


def _same_header(left: pd.DataFrame, right: pd.DataFrame) -> bool:
    return list(left.columns) == list(right.columns)


def _lines(words: list[dict], tolerance: float = 3) -> list[list[dict]]:
    lines: list[list[dict]] = []
    for word in sorted(words, key=lambda item: (round(item["top"]), item["x0"])):
        if lines and abs(lines[-1][0]["top"] - word["top"]) <= tolerance:
            lines[-1].append(word)
        else:
            lines.append([word])
    return [sorted(line, key=lambda item: item["x0"]) for line in lines]


def _phrases(line: list[dict], gap: float) -> list[tuple[float, float, str]]:
    phrases: list[list] = []
    for word in line:
        if phrases and word["x0"] - phrases[-1][1] <= gap:
            phrases[-1][1] = word["x1"]
            phrases[-1][2] += " " + word["text"]
        else:
            phrases.append([word["x0"], word["x1"], word["text"]])
    return [tuple(phrase) for phrase in phrases]


def _words_table(page, bbox) -> list[list[str]] | None:
    """Rebuild a table that has only horizontal rules from word positions."""
    lines = _lines(page.within_bbox(bbox).extract_words(keep_blank_chars=False))
    if len(lines) < 3:
        return None
    header, body = lines[0], lines[1:]
    gap = 4.0
    bands: list[list[float]] = []
    for line in body:
        for x0, x1, _ in _phrases(line, gap):
            overlapping = [band for band in bands if x0 <= band[1] and x1 >= band[0]]
            for band in overlapping:
                bands.remove(band)
                x0, x1 = min(x0, band[0]), max(x1, band[1])
            bands.append([x0, x1])
    bands.sort()
    if len(bands) < 2:
        return None

    def column(x0: float, x1: float) -> int:
        middle = (x0 + x1) / 2
        overlaps = [max(0.0, min(x1, band[1]) - max(x0, band[0])) for band in bands]
        if max(overlaps) > 0:
            return overlaps.index(max(overlaps))
        return min(range(len(bands)), key=lambda index: abs(middle - (bands[index][0] + bands[index][1]) / 2))

    rows = []
    for line in [header, *body]:
        cells = [""] * len(bands)
        for word in line:
            index = column(word["x0"], word["x1"])
            cells[index] = f"{cells[index]} {word['text']}".strip()
        rows.append(cells)
    return rows


def _overlaps(left, right) -> bool:
    return left[0] < right[2] and right[0] < left[2] and left[1] < right[3] and right[1] < left[3]


def _rule_tables(page, taken: list) -> list[tuple[float, list[list[str]]]]:
    """Find tables drawn only with horizontal rules (no vertical lines or cell fills)."""
    spans: dict[int, list[list[float]]] = {}
    for edge in page.horizontal_edges:
        row = spans.setdefault(round(edge["top"]), [])
        row.append([edge["x0"], edge["x1"]])
    rules: dict[tuple[int, int], list[float]] = {}
    for top, segments in spans.items():
        merged: list[list[float]] = []
        for x0, x1 in sorted(segments):
            if merged and x0 <= merged[-1][1] + 2:
                merged[-1][1] = max(merged[-1][1], x1)
            else:
                merged.append([x0, x1])
        for x0, x1 in merged:
            if x1 - x0 >= 60:
                rules.setdefault((round(x0), round(x1)), []).append(top)
    tables = []
    for (x0, x1), tops in rules.items():
        tops = sorted(tops)
        runs = [[tops[0]]]
        for top in tops[1:]:
            if top - runs[-1][-1] <= 40:
                runs[-1].append(top)
            else:
                runs.append([top])
        for run in runs:
            if len(run) < 2:
                continue
            spacing = sorted(b - a for a, b in zip(run, run[1:]))[(len(run) - 1) // 2]
            # The header usually sits just above the first rule.
            bbox = (x0, max(0, run[0] - spacing), x1, run[-1] + 1)
            if any(_overlaps(bbox, other) for other in taken):
                continue
            rows = _words_table(page, bbox)
            if rows:
                taken.append(bbox)
                tables.append((bbox[1], rows))
    return tables


def _page_tables(page) -> list[list[list]]:
    tables, taken = [], []
    for found in page.find_tables():
        taken.append(found.bbox)
        rows = found.extract()
        if rows and max(len(row) for row in rows) < 2:
            rows = _words_table(page, found.bbox)
        if rows:
            tables.append((found.bbox[1], rows))
    tables.extend(_rule_tables(page, taken))
    return [rows for _, rows in sorted(tables, key=lambda item: item[0])]


def pdf_to_workbook(content: bytes, file_name: str) -> bytes:
    """Extract every ruled or aligned table from a PDF into one worksheet per table."""
    import pdfplumber

    if len(content) > MAX_EXCEL_BYTES:
        raise ValueError(f"Upload a PDF no larger than {MAX_UPLOAD_MB} MB.")
    try:
        document = pdfplumber.open(BytesIO(content))
    except Exception as error:  # pdfminer raises many unrelated exception types for damaged files
        raise ValueError("The PDF could not be opened. Check that it is a valid, unencrypted PDF.") from error
    tables: list[pd.DataFrame] = []
    last_page = None
    with document:
        if len(document.pages) > MAX_PDF_PAGES:
            raise ValueError(f"A PDF can have at most {MAX_PDF_PAGES} pages for table import.")
        for page in document.pages:
            for position, rows in enumerate(_page_tables(page)):
                frame = _frame(rows)
                if frame is None:
                    continue
                # A table continued from the previous page repeats its header as the page's first table.
                if position == 0 and tables and last_page == page.page_number - 1 and _same_header(tables[-1], frame):
                    tables[-1] = pd.concat([tables[-1], frame], ignore_index=True)
                else:
                    tables.append(frame)
                last_page = page.page_number
            if len(tables) > MAX_PDF_TABLES:
                raise ValueError(f"The PDF has more than {MAX_PDF_TABLES} tables. Split it and upload the pages you need.")
    if not tables:
        raise ValueError(
            "No tables were found in this PDF. Table import needs text-based tables; scanned images are not supported."
        )
    if sum(frame.size for frame in tables) > MAX_EXCEL_CELLS:
        raise ValueError(f"The PDF tables have more than {MAX_EXCEL_CELLS:,} cells. Upload fewer pages.")
    stem = re.sub(r"[\[\]:*?/\\]", "_", file_name.rsplit(".", 1)[0]).strip().strip("'")[:24] or "PDF"
    sheets = [(stem if len(tables) == 1 else f"{stem} T{index}", frame) for index, frame in enumerate(tables, 1)]
    return generate_excel_download(sheets)
