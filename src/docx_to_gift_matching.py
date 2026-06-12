from __future__ import annotations

import argparse
import html
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NAMESPACE}


@dataclass
class MatchingQuestion:
    prompt: str
    answers: list[str]


def normalize_text(value: str) -> str:
    return " ".join(value.replace("\xa0", " ").split())


def get_text(element: ET.Element) -> str:
    parts: list[str] = []
    for text_node in element.findall('.//w:t', NS):
        parts.append(text_node.text or "")
    return normalize_text("".join(parts))


def get_text_and_stripped(element: ET.Element) -> tuple[str, str]:
    joined = "".join(text_node.text or "" for text_node in element.findall('.//w:t', NS)).strip()
    return joined, joined


def remove_bom(value: str) -> str:
    """Remove all BOM (\\ufeff) characters from the string."""
    return value.replace('\ufeff', '')


def split_merged_answers(text: str) -> list[str]:
    """
    Split a single string that may contain multiple merged answers
    like "1) First 2) Second 3) Third" into individual answer strings.

    Also handles:
    - Orphan bracket: ") text" at the start
    - No space after number: "1)text"
    - Mixed separators
    """
    # Remove all BOM characters first
    text = remove_bom(text)

    # Remove trailing "Ответ" or "Ответ: ..." or "Ответ\: ..."
    text = re.sub(r'\s*Ответ\s*[\\:]?\s*\d*\s*$', '', text, flags=re.IGNORECASE)

    # Pattern to find answer starts: "N)" or ") " (orphan bracket)
    # We split on these boundaries
    parts = re.split(r'(?=\d+\)|(?<=^)\) )', text)

    results = []
    for part in parts:
        part = part.strip()
        if not part:
            continue

        # Remove leading number + ) pattern
        clean = re.sub(r'^\d+\)\s*', '', part)
        # Handle orphan bracket at start: ") text"
        clean = re.sub(r'^\)\s+', '', clean)

        if clean:
            results.append(clean)

    return results


def is_numbered_paragraph(element: ET.Element) -> bool:
    return element.tag == f"{{{WORD_NAMESPACE}}}p" and element.find('./w:pPr/w:numPr', NS) is not None


def get_second_column_answers(table: ET.Element) -> list[str]:
    rows = table.findall('.//w:tr', NS)
    if not rows:
        return []

    first_row_cells = rows[0].findall('./w:tc', NS)
    if len(first_row_cells) < 2:
        return []

    second_cell = first_row_cells[1]
    answers: list[str] = []

    for paragraph in second_cell.findall('./w:p', NS):
        text, stripped = get_text_and_stripped(paragraph)
        if not text:
            continue

        # Remove BOM and check for answer/Ответ line
        cleaned = remove_bom(stripped)
        if cleaned.lower().startswith('ответ'):
            continue

        # Split merged answers within this paragraph
        split_answers = split_merged_answers(cleaned)
        for ans in split_answers:
            normalized = normalize_text(ans)
            if normalized:
                answers.append(normalized)

    return answers


def parse_matching_questions(docx_path: Path) -> list[MatchingQuestion]:
    if not docx_path.is_file():
        raise FileNotFoundError(f"Input file not found: {docx_path}")

    with zipfile.ZipFile(docx_path, 'r') as archive:
        document_xml = archive.read('word/document.xml')

    root = ET.fromstring(document_xml)
    body = root.find('w:body', NS)
    if body is None:
        raise ValueError('The DOCX file does not contain word/document.xml body content.')

    items = list(body)
    questions: list[MatchingQuestion] = []

    for index, element in enumerate(items):
        if not is_numbered_paragraph(element):
            continue

        prompt = get_text(element)
        if not prompt:
            continue

        if index + 1 >= len(items) or items[index + 1].tag != f"{{{WORD_NAMESPACE}}}tbl":
            continue

        answers = get_second_column_answers(items[index + 1])
        if not answers:
            continue

        questions.append(MatchingQuestion(prompt=prompt, answers=answers))

    return questions


def escape_gift(value: str) -> str:
    escaped = html.escape(value, quote=False)
    for source, replacement in {
        '{': '\\{',
        '}': '\\}',
        '=': '\\=',
        '#': '\\#',
        '~': '\\~',
        ':': '\\:',
    }.items():
        escaped = escaped.replace(source, replacement)
    return escaped


def question_to_gift(question: MatchingQuestion, index: int) -> str:
    lines = [f"::{index:02d}::[html]{escape_gift(question.prompt)} {{"]
    for answer_index, answer in enumerate(question.answers, start=1):
        lines.append(f"={escape_gift(answer)} -> {answer_index}")
    lines.append('}')
    return '\n'.join(lines)


def build_output_path(input_path: Path) -> Path:
    return input_path.with_suffix('.gift')


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description='Convert a DOCX file with numbered prompts and answer tables into GIFT matching questions.'
    )
    parser.add_argument('input', type=Path, help='Path to the source DOCX file')
    parser.add_argument(
        '-o',
        '--output',
        type=Path,
        help='Path to save the GIFT file. Defaults to the input name with .gift extension.',
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()
    output_path = args.output.resolve() if args.output else build_output_path(input_path)

    questions = parse_matching_questions(input_path)
    if not questions:
        raise ValueError('No matching questions were found in the DOCX file.')

    gift_content = '\n\n'.join(
        question_to_gift(question, index)
        for index, question in enumerate(questions, start=1)
    )
    output_path.write_text(gift_content + '\n', encoding='utf-8')

    print(f'Generated {len(questions)} questions: {output_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
