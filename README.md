# JOA AI Disclosure Scanner

A Python tool that scans PDF papers across ScienceDirect subfolders, extracts article history dates and AI declaration status, and generates a formatted Excel report.

## Features

- Scans all PDF files in ScienceDirect paper folders
- Extracts publication dates from Article History sections
- Detects AI declarations using multi-tier keyword matching
- Generates formatted Excel reports with color-coded results
- Handles fuzzy text matching for robust AI detection

## Requirements

- Python 3.8+
- pdfplumber
- openpyxl
- rapidfuzz

## Installation

```bash
pip install pdfplumber openpyxl rapidfuzz
```

## Usage

```bash
python finder.py
```

## Configuration

Edit the `ROOT_FOLDER` variable in `finder.py` to point to your ScienceDirect papers directory:

```python
ROOT_FOLDER = r"path\to\ScienceDirect\folder"
```

## Output

The script generates `AI_Disclosure_Results.xlsx` with:
- Paper filename
- Publication date
- AI declaration status (Declared, Not Found, Ambiguous)
- AI description (if provided)

## License

MIT
