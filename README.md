# PDF Page Extractor + Thumbnail GUI

A simple GUI tool for extracting specified pages from a PDF file, with thumbnails for visual page selection.


## Installation

Install dependencies:

```bash
pip install -r requirements.txt
```

Ensure Poppler is installed and accessible in your system path.

Run the app:
```bash
python pdf_extractor.pyw
```

## Features

- Thumbnail preview of all PDF pages
- Click-to-select page interface
- Manual input of page ranges (e.g., `1,3-5`)
- Saves selected pages into a new PDF
- Keeps original metadata (if any)

## Requirements

- Python 3.8+
- Poppler (required for `pdf2image`)
- See [Installation](#installation) for details
