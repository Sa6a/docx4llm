"""Public API for processing DOCX files to apply numbering."""

from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING

from lxml import etree

from docx4llm import xml_utils
from docx4llm.constants import DOCX_DOCUMENT_PART
from docx4llm.document_modifier import (
    DocumentCleaner,
    ParagraphFormatter,
    add_numbering_prefix_to_paragraph,
)
from docx4llm.docx_io import DocxFileHandler
from docx4llm.exceptions import DocxProcessingError
from docx4llm.numbering_formats import NumberingFormatterService
from docx4llm.xml_parser import NumberingParser

if TYPE_CHECKING:
    pass


class DocxNumberingProcessor:
    """Orchestrates the DOCX numbering application process."""

    def __init__(self) -> None:
        """Initialize the processor."""
        self._formatter_service = NumberingFormatterService()
        self._numbering_parser = NumberingParser(self._formatter_service)

    def process_docx(
        self,
        input_docx_path: str | Path,
        output_docx_path: str | Path,
    ) -> bool:
        """Process a DOCX file to apply numbering and save the result.

        Args:
            input_docx_path: Path to the input .docx file.
            output_docx_path: Path to save the modified .docx file.

        Returns:
            True if processing was successful, False otherwise.
            (Note: Pythonic way is to raise exceptions on failure).

        Raises:
            DocxProcessingError: If any error occurs during processing.
                                 (Replaces returning False for better error handling)

        """
        try:
            with DocxFileHandler(input_docx_path) as docx_handler:
                self._process_files_in_handler(docx_handler)
                docx_handler.create_output_docx(output_docx_path)
            return True
        except DocxProcessingError:
            raise
        except Exception as exc:
            raise DocxProcessingError(
                f"An unexpected error occurred during DOCX processing: {exc!s}",
            ) from exc

    def _process_files_in_handler(self, docx_handler: DocxFileHandler) -> None:
        """Internal logic to process XML files from the unpacked DOCX."""
        numbering_xml_content = docx_handler.get_numbering_xml_content()
        if numbering_xml_content:
            self._numbering_parser.parse_numbering_xml(numbering_xml_content)

        document_xml_content = docx_handler.get_document_xml_content()
        if not document_xml_content:
            raise DocxProcessingError(
                f"Core part {DOCX_DOCUMENT_PART} is missing or unreadable.",
            )

        modified_document_bytes = self._process_document_xml(
            document_xml_content,
        )
        docx_handler.write_document_xml_content(modified_document_bytes)

    def _process_document_xml(self, document_xml_content: bytes) -> bytes:
        """Modify the document.xml content.

        Args:
            document_xml_content: Byte content of the original document.xml.

        Returns:
            Byte content of the modified document.xml.

        """
        document_root = etree.fromstring(document_xml_content)

        DocumentCleaner.clean_numbering_references(document_root)

        paragraph_formatter = ParagraphFormatter(
            self._numbering_parser.numbering_definitions,
        )

        for paragraph_elem in xml_utils.find_all_elements(
            document_root,
            "//w:p",
        ):
            num_prefix = paragraph_formatter.format_paragraph(paragraph_elem)
            if num_prefix:
                add_numbering_prefix_to_paragraph(paragraph_elem, num_prefix)

            DocumentCleaner.remove_all_numpr_tags(paragraph_elem)

        return etree.tostring(
            document_root,
            encoding="utf-8",
            xml_declaration=True,
        )


def add_numbering_to_docx(
    input_docx_path: str | Path,
    output_docx_path: str | Path,
) -> bool:
    """Applies numbering to a DOCX file by converting numbering fields to text.

    This function maintains the public API of the original script.
    For more robust error handling, consider using try/except with
    `DocxNumberingProcessor.process_docx` and catching `DocxProcessingError`.

    Args:
        input_docx_path: Path to the source .docx file.
        output_docx_path: Path where the processed .docx file will be saved.

    Returns:
        True if the processing was successful, False if an error occurred.

    """
    processor = DocxNumberingProcessor()
    try:
        processor.process_docx(input_docx_path, output_docx_path)
        return True
    except DocxProcessingError:
        return False
