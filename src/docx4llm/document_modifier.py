"""Modifies DOCX document.xml content for numbering."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx4llm import xml_utils

if TYPE_CHECKING:
    from lxml.etree import _Element as EtreeElement

    from docx4llm.numbering_domain import NumberingDefinition
    from docx4llm.xml_parser import parse_numpr_info


class ParagraphFormatter:
    """Formats paragraphs by applying numbering prefixes."""

    def __init__(
        self,
        numbering_definitions: dict[str, NumberingDefinition],
    ) -> None:
        """Initialize the formatter.

        Args:
            numbering_definitions: A dictionary of available numbering
                                   definitions, keyed by numId.
        """
        self._numbering_definitions = numbering_definitions
        self._last_active_levels: dict[str, str] = {}

    def format_paragraph(self, paragraph_element: EtreeElement) -> str | None:
        """Formats a single paragraph element, returning its number prefix.

        This method updates the state of numbering levels based on the
        paragraph's numId and ilvl.

        Args:
            paragraph_element: The <w:p> lxml element.

        Returns:
            The numbering prefix string (e.g., "1. ", "a) ") if numbering
            is applied, otherwise None.
        """
        paragraph_properties = xml_utils.find_element(
            paragraph_element,
            ".//w:pPr",
        )
        if paragraph_properties is None:
            return None

        num_pr_element = xml_utils.find_element(
            paragraph_properties,
            ".//w:numPr",
        )

        from .xml_parser import parse_numpr_info  # noqa: PLC0415

        ilvl_str, num_id_str = parse_numpr_info(num_pr_element)

        if not self._is_valid_numbering_info(ilvl_str, num_id_str):
            return None

        ilvl, num_id = ilvl_str, num_id_str  # type: ignore[assignment]

        num_def = self._numbering_definitions[num_id]
        self._update_numbering_state(num_id, ilvl, num_def)
        return num_def.get_formatted_number(ilvl)

    def _is_valid_numbering_info(
        self,
        ilvl: str | None,
        num_id: str | None,
    ) -> bool:
        """Check if ilvl and numId constitute valid numbering information."""
        if not ilvl or not num_id:
            return False
        num_def = self._numbering_definitions.get(num_id)
        if not num_def or ilvl not in num_def.levels:
            return False
        return True

    def _update_numbering_state(
        self,
        num_id: str,
        current_ilvl_str: str,
        num_def: NumberingDefinition,
    ) -> None:
        """Update numbering level values based on current and previous state."""
        last_ilvl_str = self._last_active_levels.get(num_id)

        if last_ilvl_str is not None:
            try:
                current_level_val = int(current_ilvl_str)
                last_level_val = int(last_ilvl_str)
            except ValueError:
                self._last_active_levels[num_id] = current_ilvl_str
                return

            target_level = num_def.levels[current_ilvl_str]

            if current_level_val > last_level_val:
                pass
            elif current_level_val < last_level_val:
                target_level.increment()
                num_def.reset_levels_below(current_ilvl_str)
            else:
                target_level.increment()

        self._last_active_levels[num_id] = current_ilvl_str


class DocumentCleaner:
    """Cleans elements from a document.xml structure."""

    @staticmethod
    def clean_numbering_references(document_root: EtreeElement) -> None:
        """Remove specific w:numPr tags related to tracked changes or deletions.

        Args:
            document_root: The root element of the document.xml structure.
        """

        xpath_to_remove = (
            "//w:pPrChange//w:numPr | "
            "//w:rPrChange//w:numPr | "
            "//w:p[@w:rsidDel]//w:numPr"
        )
        for num_pr_element in xml_utils.find_all_elements(
            document_root,
            xpath_to_remove,
        ):
            parent = num_pr_element.getparent()
            if parent is not None:
                parent.remove(num_pr_element)

    @staticmethod
    def remove_all_numpr_tags(paragraph_element: EtreeElement) -> None:
        """Remove all w:numPr tags from a given paragraph element.

        Args:
            paragraph_element: The <w:p> lxml element.
        """
        for num_pr_tag in xml_utils.find_all_elements(
            paragraph_element,
            ".//w:numPr",
        ):
            parent = num_pr_tag.getparent()
            if parent is not None:
                parent.remove(num_pr_tag)


def add_numbering_prefix_to_paragraph(
    paragraph_element: EtreeElement,
    num_prefix: str,
) -> None:
    """Adds a numbering prefix string to the beginning of a paragraph.

    Args:
        paragraph_element: The <w:p> lxml element.
        num_prefix: The numbering string to prepend (e.g., "1. ").
    """
    first_text_run_text_element = xml_utils.find_element(
        paragraph_element,
        ".//w:r/w:t",
    )

    if first_text_run_text_element is not None:
        current_text = first_text_run_text_element.text or ""

        prefix_with_space = f"{num_prefix} "
        first_text_run_text_element.text = f"{prefix_with_space}{current_text}"
    else:

        run_element = xml_utils.create_element("r")
        text_element = xml_utils.create_element("t")
        text_element.text = f"{num_prefix} "
        run_element.append(text_element)

        p_pr_element = xml_utils.find_element(paragraph_element, "./w:pPr")
        if p_pr_element is not None:
            p_pr_element.addnext(run_element)
        else:
            paragraph_element.insert(0, run_element)
