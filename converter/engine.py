"""In-memory PDF-to-DOCX conversion helpers."""

from __future__ import annotations

from io import BytesIO
import logging

from docx import Document
from docx.shared import Inches
import pymupdf

from .exceptions import ConversionFailedError, EncryptedPDFError, InvalidPDFError, PDFConversionError


logger = logging.getLogger(__name__)


def convert_pdf_to_docx(pdf_bytes: bytes, preserve_images: bool = True) -> bytes:
    """Convert PDF data to a DOCX document without writing temporary files.

    Text is emitted as paragraphs in reading order. When ``preserve_images`` is
    true, each unique embedded image reference on a page is added once. Image
    failures are non-fatal because a PDF can contain image formats unsupported
    by Word; such failures are logged with their traceback.

    Raises:
        InvalidPDFError: The input is empty, malformed, or has no pages.
        EncryptedPDFError: The input requires a password.
        ConversionFailedError: Conversion could not be completed.
    """
    _validate_input(pdf_bytes)

    pdf_document = _open_pdf(pdf_bytes)
    try:
        if pdf_document.is_encrypted or pdf_document.needs_pass:
            logger.warning("Refusing encrypted PDF conversion")
            raise EncryptedPDFError("Password-protected PDFs cannot be converted.")

        if pdf_document.page_count == 0:
            logger.warning("Refusing PDF conversion because it contains no pages")
            raise InvalidPDFError("The PDF does not contain any pages.")

        word_document = Document()
        for page_index in range(pdf_document.page_count):
            page = pdf_document.load_page(page_index)
            _add_page_text(word_document, page, page_index)

            if preserve_images:
                _add_page_images(word_document, pdf_document, page, page_index)

            # A page break after the final page creates an unwanted blank page.
            if page_index < pdf_document.page_count - 1:
                word_document.add_page_break()

        with BytesIO() as output_buffer:
            word_document.save(output_buffer)
            return output_buffer.getvalue()
    except PDFConversionError:
        raise
    except Exception as exc:
        logger.exception("PDF-to-DOCX conversion failed")
        raise ConversionFailedError("Unable to convert the PDF to DOCX.") from exc
    finally:
        _close_pdf(pdf_document)


def _validate_input(pdf_bytes: bytes) -> None:
    """Reject input that cannot represent a non-empty PDF byte stream."""
    if not isinstance(pdf_bytes, bytes):
        logger.warning("PDF conversion received %s instead of bytes", type(pdf_bytes).__name__)
        raise InvalidPDFError("PDF data must be provided as bytes.")
    if not pdf_bytes:
        logger.warning("PDF conversion received empty input")
        raise InvalidPDFError("The PDF data is empty.")


def _open_pdf(pdf_bytes: bytes) -> pymupdf.Document:
    """Open *pdf_bytes* and translate reader errors into domain exceptions."""
    try:
        return pymupdf.open(stream=pdf_bytes, filetype="pdf")
    except (pymupdf.EmptyFileError, pymupdf.FileDataError, RuntimeError, ValueError) as exc:
        logger.info("Unable to open PDF input", exc_info=True)
        raise InvalidPDFError("The supplied data is not a valid PDF.") from exc
    except Exception as exc:
        logger.exception("Unexpected error while opening PDF input")
        raise ConversionFailedError("Unable to open the PDF for conversion.") from exc


def _add_page_text(word_document: Document, page: pymupdf.Page, page_index: int) -> None:
    """Append printable PDF text from one page to a Word document."""
    try:
        page_text = page.get_text("text")
    except Exception as exc:
        logger.exception("Failed to extract text from PDF page %d", page_index + 1)
        raise ConversionFailedError(
            f"Unable to extract text from page {page_index + 1}."
        ) from exc

    for line in page_text.splitlines():
        clean_line = "".join(character for character in line if character.isprintable())
        clean_line = clean_line.replace("\x00", "").strip()
        if clean_line:
            word_document.add_paragraph(clean_line)


def _add_page_images(
    word_document: Document,
    pdf_document: pymupdf.Document,
    page: pymupdf.Page,
    page_index: int,
) -> None:
    """Append each distinct embedded image on one page, logging recoverable errors."""
    try:
        image_references = page.get_images(full=True)
    except Exception:
        logger.warning(
            "Unable to list images on PDF page %d; continuing without its images",
            page_index + 1,
            exc_info=True,
        )
        return

    extracted_xrefs: set[int] = set()
    for image_reference in image_references:
        xref = image_reference[0]
        if xref in extracted_xrefs:
            continue
        extracted_xrefs.add(xref)

        try:
            image_data = pdf_document.extract_image(xref)
            image_bytes = image_data.get("image")
            if not image_bytes:
                raise ValueError("The embedded image contains no data.")

            word_document.add_picture(BytesIO(image_bytes), width=Inches(5.5))
        except Exception:
            logger.warning(
                "Unable to add embedded image %d from PDF page %d; continuing",
                xref,
                page_index + 1,
                exc_info=True,
            )


def _close_pdf(pdf_document: pymupdf.Document) -> None:
    """Close the PyMuPDF document while preserving an active conversion error."""
    try:
        pdf_document.close()
    except Exception:
        logger.exception("Failed to close PDF document")
