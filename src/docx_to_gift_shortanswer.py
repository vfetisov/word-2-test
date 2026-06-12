"""Convert Линия 1 DOCX (short answer questions) to GIFT format.

Structure expected:
  - Numbered paragraph = question text
  - Next element = table (3 rows × 2 cols: header, example, row with ?)
  - Next paragraph after table = answer line starting with "Ответ: "

The Word table is converted to an HTML table and included in the question body
(via GIFT [html] tag). Questions that contain images in their table are SKIPPED
(GIFT doesn't support images).
"""
from __future__ import annotations

import argparse
import html as html_module
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
OFFICE_DOC_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
NS = {"w": WORD_NAMESPACE}


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def remove_bom(value: str) -> str:
    """Remove all BOM (\\ufeff) characters from the string."""
    return value.replace("\ufeff", "")


def normalize_text(value: str) -> str:
    return " ".join(value.replace("\xa0", " ").split())


def get_text(element: ET.Element) -> str:
    """Concatenate all w:t text nodes inside *element*."""
    parts: list[str] = []
    for text_node in element.findall(".//w:t", NS):
        parts.append(text_node.text or "")
    return normalize_text("".join(parts))


def get_text_preserve_whitespace(element: ET.Element) -> str:
    """Like get_text but does NOT collapse whitespace via normalize_text."""
    parts: list[str] = []
    for text_node in element.findall(".//w:t", NS):
        parts.append(text_node.text or "")
    return "".join(parts).replace("\ufeff", "")


# ---------------------------------------------------------------------------
# DOCX element detection
# ---------------------------------------------------------------------------

def is_numbered_paragraph(element: ET.Element) -> bool:
    return (
        element.tag == f"{{{WORD_NAMESPACE}}}p"
        and element.find("./w:pPr/w:numPr", NS) is not None
    )


def has_images(element: ET.Element) -> bool:
    """Check if an element contains any images (blip fills) or OLE objects."""
    blip_fills = element.findall(f".//{{{DRAWINGML_NAMESPACE}}}blip")
    if blip_fills:
        return True
    ole_objects = element.findall(f".//{{{OFFICE_DOC_NAMESPACE}}}object")
    if ole_objects:
        return True
    return False


# ---------------------------------------------------------------------------
# Word table → HTML table conversion
# ---------------------------------------------------------------------------

def _cell_text(cell: ET.Element) -> str:
    """Return the plain-text content of a single table cell."""
    texts: list[str] = []
    for paragraph in cell.findall("./w:p", NS):
        t = get_text_preserve_whitespace(paragraph).strip()
        if t:
            texts.append(t)
    return " ".join(texts)


def _get_column_count(table: ET.Element) -> int:
    """Determine the number of columns from gridCol elements."""
    grid = table.find("./w:tblGrid", NS)
    if grid is not None:
        return len(grid.findall("./w:gridCol", NS))
    # fallback: use the row with the most cells
    max_cells = 0
    for row in table.findall(".//w:tr", NS):
        cells = row.findall("./w:tc", NS)
        max_cells = max(max_cells, len(cells))
    return max_cells


def table_to_html(table: ET.Element) -> str:
    """Convert a WordprocessingML table (w:tbl) into an HTML <table> string.

    Handles:
      - gridSpan (colspan)
      - vMerge (rowspan) – both "continue" and "restart"
      - simple merged cells
    """
    rows_xml = table.findall(".//w:tr", NS)
    if not rows_xml:
        return ""

    col_count = _get_column_count(table)

    # Parse all cells with their position
    # We'll build a 2D grid: grid[row][col] = cell_data or None (if merged away)
    class CellData:
        __slots__ = ("text", "colspan", "rowspan", "is_placeholder")

        def __init__(
            self,
            text: str = "",
            colspan: int = 1,
            rowspan: int = 1,
            is_placeholder: bool = False,
        ):
            self.text = text
            self.colspan = colspan
            self.rowspan = rowspan
            self.is_placeholder = is_placeholder

    grid: list[list[CellData | None]] = []

    for row_xml in rows_xml:
        cells_xml = row_xml.findall("./w:tc", NS)
        current_row: list[CellData | None] = []

        # Carry over rowspan placeholders from previous rows
        if grid:
            prev_row = grid[-1]
            for pc in prev_row:
                if pc is not None and pc.rowspan > 1 and not pc.is_placeholder:
                    # This cell continues into this row
                    current_row.append(
                        CellData(
                            text="",
                            colspan=pc.colspan,
                            rowspan=pc.rowspan - 1,
                            is_placeholder=True,
                        )
                    )
                else:
                    current_row.append(None)
        else:
            current_row = [None] * col_count

        col_idx = 0
        for cell_xml in cells_xml:
            # Skip past columns already occupied by placeholders
            while col_idx < col_count and current_row[col_idx] is not None:
                col_idx += 1
            if col_idx >= col_count:
                break

            # Determine colspan
            gridspan_el = cell_xml.find("./w:tcPr/w:gridSpan", NS)
            colspan = (
                int(gridspan_el.get(f"{{{WORD_NAMESPACE}}}val"))
                if gridspan_el is not None
                else 1
            )

            # Determine rowspan from vMerge
            vmerge_el = cell_xml.find("./w:tcPr/w:vMerge", NS)
            rowspan = 1
            if vmerge_el is not None:
                vmerge_val = vmerge_el.get(f"{{{WORD_NAMESPACE}}}val")
                if vmerge_val == "restart":
                    # Count how many subsequent rows have vMerge continue
                    current_rowspan = 1
                    for future_row in rows_xml[grid.index(row_xml) + 1:]:
                        future_cells = future_row.findall("./w:tc", NS)
                        # Find the cell at the same visual column
                        fc_idx = 0
                        found = False
                        for fc in future_cells:
                            # Skip past gridspan
                            fgs = fc.find("./w:tcPr/w:gridSpan", NS)
                            fcolspan = int(fgs.get(f"{{{WORD_NAMESPACE}}}val")) if fgs is not None else 1
                            if fc_idx <= col_idx < fc_idx + fcolspan:
                                fvm = fc.find("./w:tcPr/w:vMerge", NS)
                                if fvm is not None:
                                    fv = fvm.get(f"{{{WORD_NAMESPACE}}}val")
                                    if fv is None or fv == "continue":
                                        current_rowspan += 1
                                        found = True
                                        break
                            fc_idx += fcolspan
                            if fc_idx > col_idx:
                                break
                        if not found:
                            break
                    rowspan = current_rowspan
                # else: "continue" – handled by placeholder logic above

            cell_text = _cell_text(cell_xml)

            # Place the cell in the grid
            cell_data = CellData(
                text=cell_text, colspan=colspan, rowspan=rowspan
            )
            for c in range(colspan):
                if col_idx + c < col_count:
                    current_row[col_idx + c] = cell_data

            col_idx += colspan

        grid.append(current_row)

    # Build HTML
    html_rows: list[str] = []
    for row_idx, row in enumerate(grid):
        cells_html: list[str] = []
        col_idx = 0
        while col_idx < col_count:
            cell = row[col_idx] if col_idx < len(row) else None
            if cell is None:
                col_idx += 1
                continue
            if cell.is_placeholder:
                col_idx += cell.colspan
                continue

            tag = "th" if row_idx == 0 else "td"
            attrs = ""
            if cell.colspan > 1:
                attrs += f' colspan="{cell.colspan}"'
            if cell.rowspan > 1:
                attrs += f' rowspan="{cell.rowspan}"'

            escaped_text = html_module.escape(cell.text, quote=False)
            cells_html.append(f"<{tag}{attrs}>{escaped_text}</{tag}>")
            col_idx += cell.colspan

        html_rows.append(f"<tr>{''.join(cells_html)}</tr>")

    return f"<table>{''.join(html_rows)}</table>"


# ---------------------------------------------------------------------------
# Answer extraction – multiple correct answers
# ---------------------------------------------------------------------------

def _split_alternatives(answer_text: str) -> list[str]:
    """Split a raw answer string into a list of individual correct answers.

    Rules (from the source DOCX conventions):
      1. ``/`` separates alternatives: ``физиология/нефрология`` → [физиология, нефрология]
      2. Parenthesised content ``(…)`` is an alternative variant:
         ``генная инженерия (биотехнология)`` → [генная инженерия, биотехнология]
         BUT if the parenthesised text contains ``–`` it is a comment to discard:
         ``организменный (лишайник – это один организм!)`` → [организменный]
      3. Trailing ``.`` is stripped.
      4. Trailing/leading ``/`` and whitespace are stripped from each part.
    """
    text = answer_text.strip("/ \u00a0")

    # --- Step 1: handle parenthesised content ---
    # If parentheses contain a dash/– it's a comment → remove entirely
    text = re.sub(r"\s*\([^)]*–[^)]*\)\s*", "", text)
    # Otherwise extract parenthesised content as an alternative
    paren_match = re.search(r"\s*\(([^)]+)\)\s*", text)
    paren_alternative: str | None = None
    if paren_match:
        paren_alternative = paren_match.group(1).strip()
        text = text[: paren_match.start()].strip() + " " + text[paren_match.end():].strip()

    # --- Step 2: split by ``/`` ---
    parts = re.split(r"\s*/\s*", text)
    parts = [p.strip("/ ").strip() for p in parts if p.strip("/ ").strip()]

    # --- Step 3: add parenthesised alternative if present ---
    if paren_alternative:
        # If the parenthesised content itself contains commas, split further
        for sub in re.split(r"\s*,\s*", paren_alternative):
            sub = sub.strip().rstrip(".")
            if sub:
                parts.append(sub)

    # --- Step 4: clean up ---
    result: list[str] = []
    for p in parts:
        p = p.strip("/ \u00a0").rstrip(".")
        if p:
            result.append(p)

    # Deduplicate while preserving order
    seen: set[str] = set()
    unique: list[str] = []
    for r in result:
        if r.lower() not in seen:
            seen.add(r.lower())
            unique.append(r)
    return unique if unique else [answer_text.strip("/ ")]


def extract_answer(text: str) -> list[str] | None:
    """Extract answer(s) from a line like 'Ответ: физиология/нефрология'.

    Returns a list of individual correct answers, or ``None`` if no answer found.
    """
    text = remove_bom(text.strip())
    m = re.match(r"Ответ\s*[:\s]\s*(.+)", text, re.IGNORECASE)
    if m:
        raw = normalize_text(m.group(1))
        return _split_alternatives(raw)
    return None


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_shortanswer_questions(
    docx_path: Path,
) -> list[tuple[str, list[str]]]:
    """Parse DOCX and return list of (full_question_html, answers) tuples.

    *full_question_html* contains the prompt text followed by the HTML table.
    *answers* is a list of individual correct answer strings.
    """
    if not docx_path.is_file():
        raise FileNotFoundError(f"Input file not found: {docx_path}")

    with zipfile.ZipFile(docx_path, "r") as archive:
        document_xml = archive.read("word/document.xml")

    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("The DOCX file does not contain word/document.xml body content.")

    items = list(body)
    questions: list[tuple[str, str]] = []
    skipped_images = 0

    for index, element in enumerate(items):
        if not is_numbered_paragraph(element):
            continue

        prompt = get_text(element)
        if not prompt:
            continue

        # Next element must be a table
        if index + 1 >= len(items):
            continue
        next_tag = items[index + 1].tag
        if next_tag != f"{{{WORD_NAMESPACE}}}tbl":
            continue

        table = items[index + 1]

        # Skip if table contains images
        if has_images(table):
            skipped_images += 1
            continue

        # Convert table to HTML
        table_html = table_to_html(table)

        # Build full question: prompt + HTML table
        full_question = f"{prompt}\n\n{table_html}"

        # Find the answer paragraph after the table
        answers: list[str] | None = None
        for j in range(index + 2, min(index + 5, len(items))):
            candidate = items[j]
            if candidate.tag == f"{{{WORD_NAMESPACE}}}p":
                text = get_text(candidate)
                if not text:
                    continue
                if is_numbered_paragraph(candidate):
                    break
                ans_list = extract_answer(text)
                if ans_list:
                    answers = ans_list
                    break
            elif candidate.tag == f"{{{WORD_NAMESPACE}}}tbl":
                break

        if answers:
            questions.append((full_question, answers))

    if skipped_images:
        print(f"Skipped {skipped_images} questions containing images.")

    return questions


# ---------------------------------------------------------------------------
# GIFT output
# ---------------------------------------------------------------------------

def escape_gift(value: str) -> str:
    """Escape special GIFT characters inside a value.

    NOTE: HTML content inside [html] tags should NOT have its angle-brackets
    escaped, but GIFT special chars ({ } = # ~ :) still need escaping.
    """
    # html.escape converts & < > – we don't want that for HTML content
    # Instead we only escape the GIFT metacharacters
    for source, replacement in {
        "{": "\\{",
        "}": "\\}",
        "=": "\\=",
        "#": "\\#",
        "~": "\\~",
        ":": "\\:",
    }.items():
        value = value.replace(source, replacement)
    return value


def question_to_gift(question_html: str, answers: list[str], index: int) -> str:
    """Format one question as a GIFT short-answer block.

    Multiple correct answers are each placed on their own ``=answer`` line.
    The question body is wrapped in [html] so Moodle renders the HTML table.
    """
    lines = [f"::{index:04d}::[html]{escape_gift(question_html)} {{"]
    for ans in answers:
        lines.append(f"={escape_gift(ans)}")
    lines.append("}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def build_output_path(input_path: Path) -> Path:
    return input_path.with_suffix(".gift")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert a DOCX file with numbered prompts and answer tables "
            "into GIFT short-answer questions. Word tables are converted to "
            "HTML tables inside the question body."
        )
    )
    parser.add_argument("input", type=Path, help="Path to the source DOCX file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Path to save the GIFT file. Defaults to the input name with .gift extension.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    output_path = args.output.resolve() if args.output else build_output_path(input_path)

    questions = parse_shortanswer_questions(input_path)
    if not questions:
        raise ValueError("No questions were found in the DOCX file.")

    gift_content = "\n\n".join(
        question_to_gift(question_html, answers, index)
        for index, (question_html, answers) in enumerate(questions, start=1)
    )
    output_path.write_text(gift_content + "\n", encoding="utf-8")

    print(f"Generated {len(questions)} questions: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
