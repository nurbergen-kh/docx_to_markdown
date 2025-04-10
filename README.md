# Docs to Markdown

A Python tool to convert DOCX documents to Markdown format.

## Installation

1. Make sure you have Poetry installed:
```bash
curl -sSL https://install.python-poetry.org | python3 -
```

2. Clone the repository and install dependencies:
```bash
git clone <repository-url>
cd docs-to-markdown
poetry install
```

## Usage

To convert a DOCX file to Markdown:

```bash
poetry run python docs_to_markdown.py input.docx output.md
```

## Features

- Converts DOCX documents to clean Markdown format
- Preserves document structure (headings, tables, lists)
- Handles various table formats
- Flexible title detection
- Clean text formatting

## License

MIT 