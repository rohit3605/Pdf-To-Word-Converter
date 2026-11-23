# PDF to Word Converter

A Streamlit web application that converts PDF documents into editable Microsoft Word (DOCX) files using Python.

The converter extracts text from PDF pages and embeds extracted images into the generated Word document by default.

## Features

- Extracts text from PDF pages
- Converts extracted content into editable Word documents
- Extracts and embeds images from PDFs
- Processes PDF files directly in memory
- Handles invalid, encrypted, and empty PDFs
- Provides specific conversion error handling
- Separates the conversion engine from the Streamlit UI
- Includes automated tests for the conversion engine

## Tech Stack

- Python
- Streamlit
- PyMuPDF
- python-docx
- pytest

## Project Structure

    .
    ├── app2.py
    ├── converter/
    │   ├── __init__.py
    │   ├── engine.py
    │   └── exceptions.py
    └── tests/
        └── test_engine.py

## Implementation Improvements

This branch introduces improvements to the original conversion workflow:

- **Modular architecture:** Separated PDF-to-DOCX conversion logic from the Streamlit UI into a dedicated conversion module.
- **In-memory processing:** Replaced the temporary PDF file workflow with byte-based processing to avoid unnecessary disk writes.
- **Error handling:** Introduced dedicated exceptions for invalid PDFs, encrypted PDFs, and conversion failures.
- **Image handling:** Added image extraction and embedding support with a mechanism to handle recoverable image extraction errors.
- **Testability:** Added automated tests for conversion behavior, input validation, image handling, and failure scenarios.
- **UI integration:** Updated the Streamlit application to use the new conversion engine and provide meaningful error messages.

## How It Works

1. Upload a PDF document.
2. Click the conversion button.
3. The conversion engine processes the uploaded PDF in memory.
4. Text is extracted page by page.
5. Embedded images are extracted and added to the Word document.
6. The generated DOCX file is provided for download.

## Limitations

This converter focuses on extracting content, not reproducing the original PDF layout.

- Original formatting, fonts, colors, and positioning are not preserved.
- Extracted images are not placed at their original positions.
- Complex tables and multi-column layouts may not retain their original structure.
- Scanned PDFs without an embedded text layer are not OCR-processed.
- The resulting Word document may differ significantly from the original PDF visually.

## Setup and Installation

### 1. Clone the repository

    git clone https://github.com/rohit3605/Pdf-To-Word-Converter.git
    cd Pdf-To-Word-Converter

### 2. Create a virtual environment

Windows PowerShell:

    python -m venv .venv
    .\.venv\Scripts\Activate.ps1

### 3. Install dependencies

    pip install -r requirements.txt

If pytest is not included in the installed dependencies:

    pip install pytest

## Run the Application

Start the Streamlit app:

    streamlit run app2.py

Open the local URL provided by Streamlit in your browser.

## Run Tests

Execute the automated test suite:

    python -m pytest -q

To run tests using the project's virtual environment directly on Windows:

    .\.venv\Scripts\python.exe -m pytest -q

## Disclaimer

The generated Word document provides editable extracted content. The application is not designed to reproduce the original PDF's exact layout or formatting.