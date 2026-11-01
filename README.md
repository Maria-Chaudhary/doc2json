# Doc2JSON

**Universal Document → JSON Extractor**

Turn almost any document into clean, structured, schema-validated JSON using Google Gemini, through a polished Gradio web interface.

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Gradio](https://img.shields.io/badge/UI-Gradio-F97316)
![Gemini](https://img.shields.io/badge/LLM-Google%20Gemini-4285F4?logo=google&logoColor=white)
![Pydantic](https://img.shields.io/badge/Validation-Pydantic-E92063)
![License](https://img.shields.io/badge/License-MIT-green)

<!-- Add a screenshot of the app here: -->
<!-- ![App screenshot](docs/screenshot.png) -->

## Overview

Documents come in many formats, and getting reliable structured data out of them usually means writing a custom parser for each one. This project removes that step. Upload a file, optionally tell the model what to focus on, and get back a consistent JSON object containing a summary, key facts, entities, and sections.

Every response from the model is validated against a Pydantic schema before it is shown, so the output you download is predictable and safe to feed into other tools.

## Features

- **Wide format support:** PDF, Word (`.docx`), PowerPoint (`.pptx`), Excel (`.xlsx`, `.xlsm`), RTF, HTML, and many text and source-code formats (Markdown, CSV, JSON, YAML, Python, JavaScript, SQL, and more).
- **Structured output:** document type, title, summary, language, key information, sections, entities, and warnings.
- **Schema validation:** Gemini is constrained to a JSON schema, and the result is re-validated with Pydantic.
- **Custom instructions:** guide the extraction with a short prompt, for example "Extract all dates, amounts, and parties involved."
- **Reliable by design:** automatic retries with backoff for temporary server errors, plus a fallback model if the main one is overloaded.
- **Friendly error handling:** clear messages for password-protected PDFs, scanned PDFs, oversized files, bad API keys, and rate limits.
- **Multi-view results:** overview, key-information table, entities table, section viewer, and raw JSON with a download button.
- **Modern interface:** a custom dark "Midnight Aurora" theme built with Gradio.

## How it works

1. **Upload:** the app reads the file and extracts its text with a format-specific reader (PyMuPDF, python-docx, python-pptx, openpyxl, BeautifulSoup, striprtf).
2. **Prompt:** the text and any custom instructions are combined into a prompt.
3. **Generate:** Gemini returns JSON constrained to the document schema.
4. **Validate:** the response is checked with Pydantic and shown in the interface.
5. **Export:** download the validated result as a `.json` file.

## Output format

```json
{
  "document_type": "Research paper",
  "title": "Example Title",
  "summary": "A short summary of the document.",
  "language": "English",
  "key_information": [{ "key": "Author", "value": "Jane Doe" }],
  "sections": [{ "heading": "Introduction", "content": "..." }],
  "entities": [{ "entity_type": "Organization", "value": "Example Corp" }],
  "warnings": []
}
```

## Getting started

### Prerequisites

- Python 3.10 or newer
- A Gemini API key from [Google AI Studio](https://aistudio.google.com/app/apikey)

### Installation

```bash
git clone https://github.com/<your-username>/doc2json.git
cd doc2json

python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate

pip install -r requirements.txt
```

### Configuration

Create a `.env` file in the project root:

```env
GEMINI_API_KEY=your_api_key_here
# Optional. Defaults to gemini-2.5-flash
GEMINI_MODEL=gemini-2.5-flash
```

> Never commit your `.env` file. It is already listed in `.gitignore`.

### Run

```bash
python app.py
```

Then open the local URL printed in the terminal (usually `http://127.0.0.1:7860`).

## Project structure

```
doc2json/
├── app.py            # Gradio UI, file readers, Gemini call, validation
├── prompts.py        # System prompt and prompt builder
├── schemas.py        # Pydantic models (DocumentData)
├── requirements.txt
├── .env              # Your API key (not committed)
└── README.md
```

## Limitations

- Scanned or image-only PDFs need OCR before upload, since the app extracts embedded text only.
- Documents are limited to 120,000 characters. Split larger files into parts.
- Output quality depends on the model and the clarity of the source document, so review results before relying on them.

## Privacy

Document text is sent to the Google Gemini API for processing. Avoid uploading sensitive or confidential files unless this fits your data-handling requirements.

## Tech stack

Python · Gradio · Google Gemini (`google-genai`) · Pydantic · PyMuPDF · python-docx · python-pptx · openpyxl · BeautifulSoup · striprtf

## Contributing

Issues and pull requests are welcome. For larger changes, please open an issue first to discuss what you would like to change.

## License

Released under the [MIT License](LICENSE).
