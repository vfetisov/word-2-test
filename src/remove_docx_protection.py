from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DOCUMENT_PROTECTION_TAG = f"{{{WORD_NAMESPACE}}}documentProtection"


def remove_document_protection(input_path: Path, output_path: Path) -> bool:
    """Remove w:documentProtection from word/settings.xml in a DOCX file."""
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)

        with zipfile.ZipFile(input_path, "r") as archive:
            archive.extractall(temp_path)

        settings_path = temp_path / "word" / "settings.xml"
        if not settings_path.is_file():
            raise FileNotFoundError("The DOCX file does not contain word/settings.xml")

        ET.register_namespace("w", WORD_NAMESPACE)
        tree = ET.parse(settings_path)
        root = tree.getroot()

        removed = False
        for element in list(root):
            if element.tag == DOCUMENT_PROTECTION_TAG:
                root.remove(element)
                removed = True

        tree.write(settings_path, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for file_path in temp_path.rglob("*"):
                if file_path.is_file():
                    archive.write(file_path, file_path.relative_to(temp_path))

        return removed


def build_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_unprotected{input_path.suffix}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Remove w:documentProtection from word/settings.xml in a DOCX file."
    )
    parser.add_argument("input", type=Path, help="Path to the source DOCX file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Path to save the modified DOCX file. Defaults to <name>_unprotected.docx",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Replace the source DOCX file instead of creating a new one.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input.resolve()

    if args.in_place and args.output:
        raise ValueError("Use either --in-place or --output, not both.")

    if args.in_place:
        with tempfile.NamedTemporaryFile(delete=False, suffix=input_path.suffix) as temp_file:
            temp_output = Path(temp_file.name)

        try:
            removed = remove_document_protection(input_path, temp_output)
            shutil.move(temp_output, input_path)
            output_path = input_path
        finally:
            if temp_output.exists():
                temp_output.unlink()
    else:
        output_path = args.output.resolve() if args.output else build_output_path(input_path)
        removed = remove_document_protection(input_path, output_path)

    if removed:
        print(f"Protection tag removed: {output_path}")
    else:
        print(f"Protection tag was not found. File saved: {output_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
