"""Parses numbering.xml to create NumberingDefinition objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from lxml import etree

from docx4llm import xml_utils
from docx4llm.numbering_domain import NumberingDefinition, NumberingLevel

if TYPE_CHECKING:
    from lxml.etree import _Element as EtreeElement

    from docx4llm.numbering_formats import NumberingFormatterService


class NumberingParser:
    """Parses numbering information from numbering.xml content."""

    def __init__(self, formatter_service: NumberingFormatterService) -> None:
        """Initialize the parser with a numbering formatter service.

        Args:
            formatter_service: Service to provide formatters for numbering levels.
        """
        self._formatter_service = formatter_service
        self._abstract_num_definitions: dict[str, dict[str, dict[str, Any]]] = (
            {}
        )
        self.numbering_definitions: dict[str, NumberingDefinition] = {}

    def parse_numbering_xml(self, numbering_xml_content: bytes) -> None:
        """Parse the content of numbering.xml.

        Args:
            numbering_xml_content: The byte content of numbering.xml.
        """
        numbering_root = etree.fromstring(numbering_xml_content)
        self._parse_abstract_nums(numbering_root)
        self._parse_concrete_nums(numbering_root)

    def _parse_abstract_nums(self, numbering_root: EtreeElement) -> None:
        """Parse <w:abstractNum> elements."""
        self._abstract_num_definitions.clear()
        for abstract_num_elem in xml_utils.find_all_elements(
            numbering_root,
            "//w:abstractNum",
        ):
            abstract_num_id = xml_utils.get_attribute(
                abstract_num_elem,
                "abstractNumId",
            )
            if not abstract_num_id:
                continue

            self._abstract_num_definitions[abstract_num_id] = {}
            self._parse_abstract_num_levels(
                abstract_num_elem,
                self._abstract_num_definitions[abstract_num_id],
            )

    def _parse_abstract_num_levels(
        self,
        abstract_num_elem: EtreeElement,
        num_data: dict[str, dict[str, Any]],
    ) -> None:
        """Parse <w:lvl> elements within an <w:abstractNum>."""
        for lvl_elem in xml_utils.find_all_elements(
            abstract_num_elem,
            ".//w:lvl",
        ):
            lvl_id = xml_utils.get_attribute(lvl_elem, "ilvl")
            if not lvl_id:
                continue

            num_fmt_elem = xml_utils.find_element(lvl_elem, ".//w:numFmt")
            num_fmt_val = xml_utils.get_attribute(num_fmt_elem, "val")
            num_fmt = num_fmt_val or "decimal"

            lvl_text_elem = xml_utils.find_element(lvl_elem, ".//w:lvlText")
            lvl_text_val = xml_utils.get_attribute(lvl_text_elem, "val")
            lvl_text = lvl_text_val or f"%{int(lvl_id)+1}."

            start_elem = xml_utils.find_element(lvl_elem, ".//w:start")
            start_str = xml_utils.get_attribute(start_elem, "val")
            start_val = (
                int(start_str) if start_str and start_str.isdigit() else 1
            )

            num_data[lvl_id] = {
                "format": num_fmt,
                "text": lvl_text,
                "start": start_val,
            }

    def _parse_concrete_nums(self, numbering_root: EtreeElement) -> None:
        """Parse <w:num> elements, creating NumberingDefinition instances."""
        self.numbering_definitions.clear()
        for num_elem in xml_utils.find_all_elements(
            numbering_root,
            "//w:num",
        ):
            num_id = xml_utils.get_attribute(num_elem, "numId")
            if not num_id:
                continue

            abstract_num_id_elem = xml_utils.find_element(
                num_elem,
                ".//w:abstractNumId",
            )
            abstract_num_id_val = xml_utils.get_attribute(
                abstract_num_id_elem,
                "val",
            )
            if (
                not abstract_num_id_val
                or abstract_num_id_val not in self._abstract_num_definitions
            ):
                continue

            num_def = NumberingDefinition(abstract_num_id_val)
            self._populate_num_def_levels(
                num_def,
                self._abstract_num_definitions[abstract_num_id_val],
            )
            self._apply_level_overrides(num_elem, num_def)
            self.numbering_definitions[num_id] = num_def

    def _populate_num_def_levels(
        self,
        num_def: NumberingDefinition,
        abstract_data: dict[str, dict[str, Any]],
    ) -> None:
        """Populate NumberingDefinition with levels from abstract definition."""
        for lvl_id, lvl_data in abstract_data.items():
            formatter = self._formatter_service.get_formatter(
                lvl_data["format"]
            )
            level = NumberingLevel(
                level_id=lvl_id,
                formatter=formatter,
                text_template=lvl_data["text"],
                start_value=lvl_data["start"],
            )
            num_def.add_level(level)

    def _apply_level_overrides(
        self,
        num_elem: EtreeElement,
        num_def: NumberingDefinition,
    ) -> None:
        """Apply <w:lvlOverride> settings to the NumberingDefinition."""
        for lvl_override_elem in xml_utils.find_all_elements(
            num_elem,
            ".//w:lvlOverride",
        ):
            lvl_id = xml_utils.get_attribute(lvl_override_elem, "ilvl")
            if not lvl_id or lvl_id not in num_def.levels:
                continue

            start_override_elem = xml_utils.find_element(
                lvl_override_elem,
                ".//w:startOverride",
            )
            if start_override_elem is not None:
                new_start_str = xml_utils.get_attribute(
                    start_override_elem, "val"
                )
                if new_start_str and new_start_str.isdigit():
                    new_start_val = int(new_start_str)
                    overridden_level = num_def.levels[lvl_id]
                    overridden_level.start_value = new_start_val
                    overridden_level.reset()


def parse_numpr_info(
    numpr_element: EtreeElement | None,
) -> tuple[str | None, str | None]:
    """Extract ilvl and numId from a w:numPr element.

    Args:
        numpr_element: The w:numPr lxml element.

    Returns:
        A tuple containing (ilvl string, numId string), or (None, None).
    """
    if numpr_element is None:
        return None, None

    ilvl_element = xml_utils.find_element(numpr_element, ".//w:ilvl")
    num_id_element = xml_utils.find_element(numpr_element, ".//w:numId")

    ilvl = xml_utils.get_attribute(ilvl_element, "val")
    num_id = xml_utils.get_attribute(num_id_element, "val")

    return ilvl, num_id
