class PDFConversionError(Exception):
    """Base exception for all PDF-to-Word conversion errors."""


class InvalidPDFError(PDFConversionError):
    """Raised when the uploaded file is not a valid PDF."""


class EncryptedPDFError(PDFConversionError):
    """Raised when the PDF is password-protected."""


class ConversionFailedError(PDFConversionError):
    """Raised when PDF-to-Word conversion fails unexpectedly."""
