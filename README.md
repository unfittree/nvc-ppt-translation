# PPT Translation (English -> Chinese)

A Python CLI project that translates PowerPoint `.ppt/.pptx` files from English to Chinese while keeping formatting as close as possible.

## What it does

- Translates text in slide text boxes and table cells.
- Handles grouped shapes recursively.
- Optionally translates speaker notes.
- Uses caching so repeated text is translated once.
- Applies NVC and internal-growth terminology with glossary enforcement.
- Translates paragraph text and redistributes into original runs to preserve style layout.
- Supports legacy `.ppt` input by converting to `.pptx` with local PowerPoint.
- Normalizes Chinese output to Simplified Chinese by default.

## Requirements

- Python 3.9+
- Internet connection (for Google translation service via `deep-translator`)
- Microsoft PowerPoint (required when input is `.ppt`)

## Setup

```powershell
cd C:\Users\guwei\ppt-translation
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Usage

```powershell
python -m ppt_translation.cli "C:\path\input.ppt"
```

This creates `C:\path\input_zh.pptx`.

Use a custom output path:

```powershell
python -m ppt_translation.cli "C:\path\input.pptx" -o "C:\path\input_cn.pptx"
```

Translate notes too:

```powershell
python -m ppt_translation.cli "C:\path\input.pptx" --include-notes
```

Use a custom glossary file:

```powershell
python -m ppt_translation.cli "C:\path\input.pptx" --glossary-file "C:\path\my_glossary.csv"
```

If `domain_glossary.csv` exists in the working directory, it is applied automatically.

## Optional install as command

```powershell
pip install -e .
ppt-translate "C:\path\input.pptx"
```

## Notes

- Input supports `.ppt` and `.pptx`; output is always `.pptx`.
- Some very complex text objects (for example chart embedded labels) may not be translated.
