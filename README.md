## Contents

This repository contains:
- `bookmark-enrich.py`, a script for enriching bookmarks with status checks, title extraction, favicon discovery, metadata extraction, and description generation.
- `encrypt-private-json.py`, a script to encrypt private.json files
- `Bookmarks-Test.xlsm`, a sample of the starter excel file used to generate both the publically accessible bookmarks.json and private private.json files.

## Setup

Use the included local virtual environment or recreate it before running the script.

```bash
cd '/home/sasha/Documents/Projects/bookmark-site/bookmark-site-basic__chatGPT-v1.2-Final__python-excel-file'
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
```

## Environment Variables

The script uses Google Gemini for description generation. Set the API key before running:

```bash
export GEMINI_API_KEY="your_api_key_here"
```

You can also use the provided `.env.example` as a reference.

## Running

Run the script from the project root:

```bash
.venv/bin/python bookmark-enrich.py <path-to-workbook.xlsm> [options]
```

For example:

```bash
.venv/bin/python bookmark-enrich.py Bookmarks-Test.xlsm --metadata --titles --download-favicons
```

## Notes

- Use the `.venv` Python interpreter to ensure dependencies are resolved correctly.
- `google-genai` is the supported package for Gemini API access.
