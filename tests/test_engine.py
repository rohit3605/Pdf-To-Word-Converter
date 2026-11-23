"""Focused tests for the in-memory PDF-to-DOCX engine."""

from __future__ import annotations

from io import BytesIO

from docx import Document
import pymupdf
import pytest

from converter import engine
from converter.engine import convert_pdf_to_docx
from converter.exceptions import ConversionFailedError, EncryptedPDFError, InvalidPDFError


def _make_pdf(*page_texts: str) -> bytes:
    document = pymupdf.open()
    try:
        for text in page_texts:
            page = document.new_page()
            page.insert_text((72, 72), text)
        return document.tobytes()
    finally:
        document.close()


def _open_docx(docx_bytes: bytes) -> Document:
    return Document(BytesIO(docx_bytes))


def _make_zero_page_pdf() -> bytes:
    objects = (
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
        "2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n",
    )
    pdf = "%PDF-1.4\n"
    offsets: list[int] = []
    for object_text in objects:
        offsets.append(len(pdf.encode("ascii")))
        pdf += object_text

    xref_offset = len(pdf.encode("ascii"))
    pdf += "xref\n0 3\n0000000000 65535 f \n"
    pdf += "".join(f"{offset:010d} 00000 n \n" for offset in offsets)
    pdf += f"trailer\n<< /Root 1 0 R /Size 3 >>\nstartxref\n{xref_offset}\n%%EOF\n"
    return pdf.encode("ascii")


def _make_png() -> bytes:
    pixmap = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.Rect(0, 0, 10, 10), 0)
    pixmap.set_rect(pixmap.irect, (220, 20, 60))
    return pixmap.tobytes("png")


def test_converts_valid_pdf_and_returns_valid_docx() -> None:
    result = convert_pdf_to_docx(_make_pdf("First page", "Second page"))

    document = _open_docx(result)
    paragraph_text = [paragraph.text for paragraph in document.paragraphs]

    assert result.startswith(b"PK")
    assert "First page" in paragraph_text
    assert "Second page" in paragraph_text
    assert document.element.body.xml.count('w:type="page"') == 1


def test_rejects_invalid_pdf_bytes() -> None:
    with pytest.raises(InvalidPDFError, match="valid PDF"):
        convert_pdf_to_docx(b"this is not a PDF")


def test_rejects_empty_input() -> None:
    with pytest.raises(InvalidPDFError, match="empty"):
        convert_pdf_to_docx(b"")


@pytest.mark.parametrize("pdf_bytes", [None, "not bytes"])
def test_rejects_non_bytes_input(pdf_bytes: object) -> None:
    with pytest.raises(InvalidPDFError, match="must be provided as bytes"):
        convert_pdf_to_docx(pdf_bytes)  # type: ignore[arg-type]


def test_rejects_zero_page_pdf() -> None:
    with pytest.raises(InvalidPDFError, match="does not contain any pages"):
        convert_pdf_to_docx(_make_zero_page_pdf())


def test_rejects_encrypted_pdf() -> None:
    document = pymupdf.open()
    try:
        document.new_page().insert_text((72, 72), "Secret")
        encrypted_pdf = document.tobytes(
            encryption=pymupdf.PDF_ENCRYPT_AES_256,
            owner_pw="owner-password",
            user_pw="user-password",
        )
    finally:
        document.close()

    with pytest.raises(EncryptedPDFError, match="Password-protected"):
        convert_pdf_to_docx(encrypted_pdf)


def test_preserves_images_by_default_and_deduplicates_per_page() -> None:
    pdf_document = pymupdf.open()
    try:
        page = pdf_document.new_page()
        image_xref = page.insert_image(pymupdf.Rect(72, 72, 144, 144), stream=_make_png())
        page.insert_image(pymupdf.Rect(180, 72, 252, 144), xref=image_xref)
        pdf_bytes = pdf_document.tobytes()
    finally:
        pdf_document.close()

    document = _open_docx(convert_pdf_to_docx(pdf_bytes))

    assert len(document.inline_shapes) == 1


def test_can_omit_images() -> None:
    pdf_document = pymupdf.open()
    try:
        page = pdf_document.new_page()
        page.insert_image(pymupdf.Rect(72, 72, 144, 144), stream=_make_png())
        pdf_bytes = pdf_document.tobytes()
    finally:
        pdf_document.close()

    document = _open_docx(convert_pdf_to_docx(pdf_bytes, preserve_images=False))

    assert len(document.inline_shapes) == 0


def test_image_extraction_failure_is_logged_and_conversion_continues(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    class FailingImagePage:
        def get_text(self, option: str) -> str:
            assert option == "text"
            return "Text survives image extraction failure"

        def get_images(self, *, full: bool) -> list[tuple[int]]:
            assert full is True
            return [(23,)]

    class FailingImageDocument:
        is_encrypted = False
        needs_pass = False
        page_count = 1

        def __init__(self) -> None:
            self.closed = False

        def load_page(self, page_index: int) -> FailingImagePage:
            assert page_index == 0
            return FailingImagePage()

        def extract_image(self, xref: int) -> dict[str, bytes]:
            assert xref == 23
            raise RuntimeError("broken embedded image")

        def close(self) -> None:
            self.closed = True

    pdf_document = FailingImageDocument()
    monkeypatch.setattr(engine, "_open_pdf", lambda _: pdf_document)

    with caplog.at_level("WARNING", logger=engine.__name__):
        result = convert_pdf_to_docx(_make_pdf("ignored by the fake document"))

    assert result.startswith(b"PK")
    assert pdf_document.closed is True
    assert "Unable to add embedded image 23 from PDF page 1; continuing" in caplog.text


def test_unexpected_internal_error_is_chained_as_conversion_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    original_error = RuntimeError("Word document initialization failed")

    def fail_to_create_document() -> None:
        raise original_error

    monkeypatch.setattr(engine, "Document", fail_to_create_document)

    with pytest.raises(ConversionFailedError) as raised_error:
        convert_pdf_to_docx(_make_pdf("Content"))

    assert raised_error.value.__cause__ is original_error
