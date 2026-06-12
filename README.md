# word-2-test

Independent Python project created in `Projects`, separate from `FS API`.

## Structure

- `.venv/` — local virtual environment folder
- `src/` — source code
- `requirements.txt` — project dependencies

## Getting started

Create or activate the virtual environment, then install dependencies:

### Windows (cmd)

```cmd
word-2-test\.venv\Scripts\activate
pip install -r requirements.txt
```

## Notes

The only external dependency at the moment is `PyInstaller`, which is used to build a standalone `.exe`.

## DOCX protection removal script

The project includes a script that:

- opens a `.docx` file as an archive;
- finds `word/settings.xml`;
- removes the `w:documentProtection` tag;
- saves the result back as `.docx`.

Run it like this:

```cmd
python src\remove_docx_protection.py examples\example.docx
```

By default, the script creates a new file with `_unprotected` added to the name.

Optional flags:

```cmd
python src\remove_docx_protection.py input.docx --in-place
python src\remove_docx_protection.py input.docx --output output.docx
```

## Build EXE

Install dependencies:

```cmd
pip install -r requirements.txt
```

Build the executable:

```cmd
pyinstaller --onefile --name remove_docx_protection src\remove_docx_protection.py
```

The result will appear in:

```cmd
dist\remove_docx_protection.exe
```

## DOCX to GIFT matching converter

The project also includes a script for converting a specific DOCX layout into GIFT matching questions.

Expected structure:

- each question is a numbered paragraph;
- the question text is the paragraph content;
- answers are taken only from the second column of the table immediately below the paragraph;
- the first and third columns are ignored;
- answer order in the second column defines the matching targets;
- a trailing line starting with `Ответ` in the second column is ignored.

Run it like this:

```cmd
python src\docx_to_gift_matching.py examples\Линия 12 Для теста.docx
```

Optional output path:

```cmd
python src\docx_to_gift_matching.py examples\Линия 12 Для теста.docx --output examples\Линия 12 Для теста.gift
```
