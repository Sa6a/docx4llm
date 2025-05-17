"""Custom exceptions for the package."""


class DocxNumberingError(Exception):
    """Base exception for errors in this package."""


class DocxProcessingError(DocxNumberingError):
    """Raised for errors during DOCX file processing."""


class NumberingDefinitionError(DocxNumberingError):
    """Raised for errors related to numbering definitions."""


class PandocError(DocxNumberingError):
    """Base exception for Pandoc related errors."""


class PandocNotInstalledError(PandocError):
    """Raised when pypandoc is not installed but required."""


class PandocConversionError(PandocError):
    """Raised when pypandoc fails to convert a file."""
