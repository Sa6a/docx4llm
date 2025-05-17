"""DOCX Numbering Processor Package.

Provides tools to process DOCX files and apply numbering formatting directly
into the document text.
"""

from .api import DocxNumberingProcessor, add_numbering_to_docx
from .exceptions import (
    DocxNumberingError,
    DocxProcessingError,
    NumberingDefinitionError,
    PandocConversionError,
    PandocError,
    PandocNotInstalledError,
)
from .pandoc_utils import convert_docx_to_format

__all__ = [
    "add_numbering_to_docx",
    "DocxNumberingProcessor",
    "convert_docx_to_format",
    "DocxNumberingError",
    "DocxProcessingError",
    "NumberingDefinitionError",
    "PandocError",
    "PandocNotInstalledError",
    "PandocConversionError",
]

__version__ = "0.1.0"
