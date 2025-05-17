"""Utilities for document conversion using Pandoc."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Literal

from docx4llm.exceptions import (
    PandocConversionError,
    PandocError,
    PandocNotInstalledError,
)

TrackChangesOption = Literal["accept", "reject", "all"]


def convert_docx_to_format(
    input_docx_path: str | Path,
    output_format: str,
    track_changes: TrackChangesOption = "all",
) -> str:
    """Converts a DOCX file to another format using Pandoc.

    Args:
        input_docx_path: Path to the input .docx file.
        output_format: The desired output format (e.g., "md", "html", "pdf").
        track_changes: How to handle tracked changes.
                       Options: "accept", "reject", "all". Defaults to "all".

    Returns:
        The path to the converted output file.

    Raises:
        PandocNotInstalledError: If pypandoc is not installed.
        PandocConversionError: If any error occurs during conversion.
        FileNotFoundError: If the input_docx_path does not exist.
    """
    input_path = Path(input_docx_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Input DOCX file not found: {input_path}")

    try:
        import pypandoc
    except ImportError:
        raise PandocNotInstalledError(
            "pypandoc library is not installed. "
            "Please install it with: pip install pypandoc",
        ) from None

    output_file_path = input_path.with_suffix(f".{output_format}")

    extra_args = []
    valid_track_changes: list[TrackChangesOption] = ["accept", "reject", "all"]
    if track_changes not in valid_track_changes:
        effective_track_changes: TrackChangesOption = "all"
    else:
        effective_track_changes = track_changes
    extra_args.append(f"--track-changes={effective_track_changes}")

    try:
        pypandoc.convert_file(
            str(input_path),
            output_format,
            outputfile=str(output_file_path),
            extra_args=extra_args,
        )
        return str(output_file_path)
    except Exception as e:
        raise PandocConversionError(
            f"Pandoc conversion failed for '{input_path}' to '{output_format}': {e!s}",
        ) from e
