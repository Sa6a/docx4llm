"""Handles I/O operations for DOCX files (unpacking and packing)."""

from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING


from docx4llm.constants import (
    DOCX_DOCUMENT_PART,
    DOCX_NUMBERING_PART,
    TEMP_DIR_SUFFIX,
)
from docx4llm.exceptions import DocxProcessingError

if TYPE_CHECKING:
    pass


class DocxFileHandler:
    """Manages unpacking and packing of DOCX files."""

    def __init__(self, input_docx_path: str | Path) -> None:
        """Initialize with the path to the input DOCX file.

        Args:
            input_docx_path: Path to the .docx file.
        """
        self.input_path = Path(input_docx_path)
        if not self.input_path.exists() or not self.input_path.is_file():
            raise FileNotFoundError(
                f"Input DOCX file not found: {self.input_path}"
            )
        self._temp_dir_path: Path | None = None

    def __enter__(self) -> DocxFileHandler:
        """Enter context management: unpacks the DOCX to a temporary directory."""

        self._temp_dir_path = Path(tempfile.mkdtemp(suffix=TEMP_DIR_SUFFIX))

        if self._temp_dir_path.exists():
            shutil.rmtree(self._temp_dir_path)
        self._temp_dir_path.mkdir(parents=True, exist_ok=True)

        try:
            with zipfile.ZipFile(self.input_path, "r") as zip_ref:
                zip_ref.extractall(self._temp_dir_path)
        except zipfile.BadZipFile as exc:
            self.cleanup()
            raise DocxProcessingError(
                f"Invalid DOCX file (not a valid zip archive): {self.input_path}",
            ) from exc
        except Exception as exc:
            self.cleanup()
            raise DocxProcessingError(
                f"Failed to extract DOCX file: {self.input_path}",
            ) from exc
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:  # type: ignore[no-untyped-def]
        """Exit context management: cleans up the temporary directory."""
        self.cleanup()

    def cleanup(self) -> None:
        """Remove the temporary directory if it exists."""
        if self._temp_dir_path and self._temp_dir_path.exists():
            shutil.rmtree(self._temp_dir_path, ignore_errors=True)
        self._temp_dir_path = None

    def _assert_temp_dir(self) -> Path:
        """Ensure the temporary directory is available."""
        if not self._temp_dir_path or not self._temp_dir_path.is_dir():
            raise DocxProcessingError(
                "Temporary directory not initialized or accessible.",
            )
        return self._temp_dir_path

    def get_numbering_xml_content(self) -> bytes | None:
        """Read the content of word/numbering.xml from the unpacked DOCX.

        Returns:
            Byte content of numbering.xml, or None if not found.
        """
        temp_dir = self._assert_temp_dir()
        numbering_path = temp_dir / DOCX_NUMBERING_PART
        if numbering_path.exists():
            return numbering_path.read_bytes()
        return None

    def get_document_xml_content(self) -> bytes | None:
        """Read the content of word/document.xml from the unpacked DOCX.

        Returns:
            Byte content of document.xml, or None if not found.
        """
        temp_dir = self._assert_temp_dir()
        document_path = temp_dir / DOCX_DOCUMENT_PART
        if document_path.exists():
            return document_path.read_bytes()
        raise DocxProcessingError(
            f"Required part {DOCX_DOCUMENT_PART} missing."
        )

    def write_document_xml_content(self, content: bytes) -> None:
        """Write modified content back to word/document.xml in the temp dir.

        Args:
            content: The modified byte content for document.xml.
        """
        temp_dir = self._assert_temp_dir()
        document_path = temp_dir / DOCX_DOCUMENT_PART
        document_path.write_bytes(content)

    def create_output_docx(self, output_docx_path: str | Path) -> None:
        """Create a new DOCX file from the contents of the temporary directory.

        Args:
            output_docx_path: Path to save the new .docx file.
        """
        source_dir = self._assert_temp_dir()
        output_path = Path(output_docx_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root_str, _, files in os.walk(source_dir):
                root_path = Path(root_str)
                for filename in files:
                    file_path = root_path / filename
                    arcname = file_path.relative_to(source_dir)
                    zipf.write(file_path, arcname)
