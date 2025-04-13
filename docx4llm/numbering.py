import zipfile
from lxml import etree
from typing import Dict, Tuple, List, Optional, Any


class NumberingFormat:
    """Class for formatting numbers in various styles."""

    @staticmethod
    def format_number(number: int, format_type: str) -> str:
        """Formats a number according to the specified format."""
        if format_type == "decimal":
            return str(number)
        elif format_type == "upperRoman":
            return NumberingFormat._to_roman(number)
        elif format_type == "lowerRoman":
            return NumberingFormat._to_roman(number).lower()
        elif format_type == "upperLetter":
            return chr(ord("A") + number - 1)
        elif format_type == "lowerLetter":
            return chr(ord("a") + number - 1)
        return str(number)

    @staticmethod
    def _to_roman(number: int) -> str:
        """Converts a number to Roman numerals."""
        roman_map = {
            1000: "M",
            900: "CM",
            500: "D",
            400: "CD",
            100: "C",
            90: "XC",
            50: "L",
            40: "XL",
            10: "X",
            9: "IX",
            5: "V",
            4: "IV",
            1: "I",
        }
        result = ""
        for value, numeral in roman_map.items():
            while number >= value:
                result += numeral
                number -= value
        return result


class NumberingLevel:
    """Represents one level of numbering."""

    def __init__(self, format_type: str, text_template: str, start_value: int):
        self.format_type = format_type
        self.text_template = text_template
        self.start_value = start_value
        self.current_value = start_value

    def reset(self) -> None:
        """Resets the numbering value to the initial value."""
        self.current_value = self.start_value

    def increment(self) -> None:
        """Increases the numbering value by 1."""
        self.current_value += 1

    def format_current_value(self) -> str:
        """Formats the current value according to the format type."""
        return NumberingFormat.format_number(self.current_value, self.format_type)


class NumberingDefinition:
    """Represents a numbering definition in the document."""

    def __init__(self, abstract_num_id: str):
        self.abstract_num_id = abstract_num_id
        self.levels: Dict[str, NumberingLevel] = {}

    def add_level(self, level_id: str, level: NumberingLevel) -> None:
        """Adds a numbering level."""
        self.levels[level_id] = level

    def reset_levels_below(self, current_level_id: str) -> None:
        """Resets all levels below the current one to their initial values."""
        current_level_int = int(current_level_id)
        for level_id, level in self.levels.items():
            if int(level_id) > current_level_int:
                level.reset()

    def get_formatted_number(self, level_id: str) -> str:
        """Gets the formatted number for the specified level."""
        if level_id not in self.levels:
            return ""

        level = self.levels[level_id]
        text = level.text_template

        level_ids = sorted(self.levels.keys(), key=lambda x: int(x))
        for sub_level_id in level_ids:
            sub_level = self.levels[sub_level_id]
            placeholder = f"%{int(sub_level_id) + 1}"
            if placeholder in text:
                text = text.replace(placeholder, sub_level.format_current_value())

        return text


class XmlHelper:
    """Helper class for working with Word document XML."""

    WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_MAP = {"w": WORD_NS}

    @staticmethod
    def find_element(parent, xpath: str) -> Optional[etree.Element]:
        """Finds an element by XPath."""
        return parent.find(xpath, namespaces=XmlHelper.NS_MAP)

    @staticmethod
    def find_all_elements(parent, xpath: str) -> List[etree.Element]:
        """Finds all elements by XPath."""
        return parent.xpath(xpath, namespaces=XmlHelper.NS_MAP)

    @staticmethod
    def get_attribute(element, attr_name: str) -> Optional[str]:
        """Gets an attribute value from an element."""
        if element is None:
            return None

        full_attr_name = f"{{{XmlHelper.WORD_NS}}}{attr_name}"
        return element.get(full_attr_name)

    @staticmethod
    def parse_numbering_info(numPr_element) -> Tuple[Optional[str], Optional[str]]:
        """Extracts numbering information from a numPr element."""
        if numPr_element is None:
            return None, None

        ilvl_element = XmlHelper.find_element(numPr_element, ".//w:ilvl")
        num_id_element = XmlHelper.find_element(numPr_element, ".//w:numId")

        ilvl = XmlHelper.get_attribute(ilvl_element, "val")
        num_id = XmlHelper.get_attribute(num_id_element, "val")

        return ilvl, num_id

    @staticmethod
    def has_element(parent, xpath: str) -> bool:
        """Checks for the presence of an element by XPath."""
        return len(XmlHelper.find_all_elements(parent, xpath)) > 0

    @staticmethod
    def element_exists(element, xpath: str) -> bool:
        """Checks if an element exists by XPath."""
        return XmlHelper.find_element(element, xpath) is not None


class NumberingParser:
    """Analyzes the numbering structure in a Word document."""

    def __init__(self):
        self.abstract_numbering_data: Dict[str, Dict[str, Dict[str, Any]]] = {}
        self.numbering_definitions: Dict[str, NumberingDefinition] = {}

    def parse_numbering_xml(self, numbering_xml_content: bytes) -> None:
        """Parses the content of the numbering.xml file from DOCX."""
        numbering_root = etree.fromstring(numbering_xml_content)

        self._parse_abstract_numbering(numbering_root)
        self._parse_numbering(numbering_root)

    def _parse_abstract_numbering(self, numbering_root: etree.Element) -> None:
        """Parses abstract numbering definitions."""
        for abstract_num in XmlHelper.find_all_elements(
            numbering_root, "//w:abstractNum"
        ):
            abstract_num_id = XmlHelper.get_attribute(abstract_num, "abstractNumId")
            if not abstract_num_id:
                continue

            self.abstract_numbering_data[abstract_num_id] = {}

            for lvl in XmlHelper.find_all_elements(abstract_num, ".//w:lvl"):
                ilvl = XmlHelper.get_attribute(lvl, "ilvl")
                if not ilvl:
                    continue

                num_fmt_element = XmlHelper.find_element(lvl, ".//w:numFmt")
                num_fmt = XmlHelper.get_attribute(num_fmt_element, "val") or "decimal"

                lvl_text_element = XmlHelper.find_element(lvl, ".//w:lvlText")
                lvl_text = XmlHelper.get_attribute(lvl_text_element, "val") or "%1."

                start_element = XmlHelper.find_element(lvl, ".//w:start")
                start_str = XmlHelper.get_attribute(start_element, "val")
                start = int(start_str) if start_str else 1

                self.abstract_numbering_data[abstract_num_id][ilvl] = {
                    "format": num_fmt,
                    "text": lvl_text,
                    "start": start,
                }

    def _parse_numbering(self, numbering_root: etree.Element) -> None:
        """Parses specific numbering definitions."""
        for num in XmlHelper.find_all_elements(numbering_root, "//w:num"):
            num_id = XmlHelper.get_attribute(num, "numId")
            if not num_id:
                continue

            abstract_num_id_element = XmlHelper.find_element(num, ".//w:abstractNumId")
            abstract_num_id = XmlHelper.get_attribute(abstract_num_id_element, "val")
            if (
                not abstract_num_id
                or abstract_num_id not in self.abstract_numbering_data
            ):
                continue

            num_def = NumberingDefinition(abstract_num_id)

            for lvl_id, lvl_data in self.abstract_numbering_data[
                abstract_num_id
            ].items():
                level = NumberingLevel(
                    lvl_data["format"], lvl_data["text"], lvl_data["start"]
                )
                num_def.add_level(lvl_id, level)

            for lvl_override in XmlHelper.find_all_elements(num, ".//w:lvlOverride"):
                ilvl = XmlHelper.get_attribute(lvl_override, "ilvl")
                if not ilvl or ilvl not in num_def.levels:
                    continue

                start_override = XmlHelper.find_element(
                    lvl_override, ".//w:startOverride"
                )
                if start_override:
                    new_start_str = XmlHelper.get_attribute(start_override, "val")
                    if new_start_str:
                        new_start = int(new_start_str)
                        num_def.levels[ilvl].start_value = new_start
                        num_def.levels[ilvl].current_value = new_start

            self.numbering_definitions[num_id] = num_def


class ParagraphFormatter:
    """Formats paragraphs with numbering."""

    def __init__(self, numbering_definitions: Dict[str, NumberingDefinition]):
        self.numbering_definitions = numbering_definitions
        self.last_active_levels = {}

    def format_paragraph(self, paragraph: etree.Element) -> str:
        """Formats a paragraph with numbering."""
        paragraph_text = self._get_paragraph_text(paragraph)

        pPr = XmlHelper.find_element(paragraph, ".//w:pPr")
        if pPr is None:
            return paragraph_text

        numPr = XmlHelper.find_element(pPr, ".//w:numPr")
        ilvl, num_id = XmlHelper.parse_numbering_info(numPr)

        num_prefix = ""
        if self._has_valid_numbering(ilvl, num_id):
            num_def = self.numbering_definitions[num_id]

            last_ilvl = self.last_active_levels.get(num_id)

            if last_ilvl is not None:
                current_level = int(ilvl)
                last_level = int(last_ilvl)

                if current_level > last_level:
                    pass
                elif current_level < last_level:
                    num_def.levels[ilvl].increment()
                    num_def.reset_levels_below(ilvl)
                else:
                    num_def.levels[ilvl].increment()

            num_prefix = num_def.get_formatted_number(ilvl)

            self.last_active_levels[num_id] = ilvl

        if num_prefix:
            return f"{num_prefix} {paragraph_text}"
        return paragraph_text

    def _has_valid_numbering(self, ilvl: Optional[str], num_id: Optional[str]) -> bool:
        """Checks if a paragraph has valid numbering."""
        if not ilvl or not num_id or num_id not in self.numbering_definitions:
            return False

        if ilvl not in self.numbering_definitions[num_id].levels:
            return False

        return True

    def _get_paragraph_text(self, paragraph: etree.Element) -> str:
        """Extracts text from a paragraph."""
        text_parts = []
        for t in XmlHelper.find_all_elements(paragraph, ".//w:t"):
            if t.text is not None:
                text_parts.append(t.text)
        return "".join(text_parts)


class DocumentCleaner:
    """Class for preliminary document cleaning."""

    @staticmethod
    def clean_document(document_root: etree.Element) -> None:
        """
        Cleans the document according to requirements:
        1. Removes paragraphs with w:numPr in w:pPrChange without w:t
        2. Removes all w:pPrChange and w:rPrChange tags
        """
        DocumentCleaner._remove_empty_numbered_paragraphs(document_root)
        DocumentCleaner._remove_change_tags(document_root)

    @staticmethod
    def _remove_empty_numbered_paragraphs(document_root: etree.Element) -> None:
        """Removes empty paragraphs with numbering inside change tags."""
        paragraphs_to_remove = []

        for paragraph in XmlHelper.find_all_elements(document_root, "//w:p"):
            has_text = XmlHelper.element_exists(paragraph, ".//w:t")
            has_numPr_in_change = XmlHelper.element_exists(
                paragraph, ".//w:pPrChange//w:numPr"
            )

            if not has_text and has_numPr_in_change:
                paragraphs_to_remove.append(paragraph)

        for paragraph in paragraphs_to_remove:
            parent = paragraph.getparent()
            if parent is not None:
                parent.remove(paragraph)

    @staticmethod
    def _remove_change_tags(document_root: etree.Element) -> None:
        """Removes all change tags."""
        for change_tag in XmlHelper.find_all_elements(document_root, "//w:pPrChange"):
            parent = change_tag.getparent()
            if parent is not None:
                parent.remove(change_tag)

        for change_tag in XmlHelper.find_all_elements(document_root, "//w:rPrChange"):
            parent = change_tag.getparent()
            if parent is not None:
                parent.remove(change_tag)


class DocxToTextConverter:
    """Converts DOCX files to TXT while preserving numbering."""

    def __init__(self):
        self.numbering_parser = NumberingParser()

    def convert(self, docx_path: str, txt_path: str) -> bool:
        """Converts a DOCX file to TXT while preserving numbering."""
        try:
            with zipfile.ZipFile(docx_path, "r") as docx:
                self._parse_numbering(docx)
                return self._process_document(docx, txt_path)
        except Exception as e:
            print(f"Error converting file: {e}")
            return False

    def _parse_numbering(self, docx: zipfile.ZipFile) -> None:
        """Parses the numbering file if it exists."""
        try:
            numbering_xml = docx.read("word/numbering.xml")
            self.numbering_parser.parse_numbering_xml(numbering_xml)
        except KeyError:
            print("File word/numbering.xml not found.")

    def _process_document(self, docx: zipfile.ZipFile, txt_path: str) -> bool:
        """Processes the main document and writes the result."""
        try:
            document_xml = docx.read("word/document.xml")
            document_root = etree.fromstring(document_xml)

            DocumentCleaner.clean_document(document_root)

            paragraph_formatter = ParagraphFormatter(
                self.numbering_parser.numbering_definitions
            )

            output_lines = []
            for paragraph in XmlHelper.find_all_elements(document_root, "//w:p"):
                formatted_paragraph = paragraph_formatter.format_paragraph(paragraph)
                if formatted_paragraph.strip():
                    output_lines.append(formatted_paragraph)

            with open(txt_path, "w", encoding="utf-8") as outfile:
                for line in output_lines:
                    outfile.write(line + "\n")

            return True

        except KeyError:
            print("File word/document.xml not found.")
            return False


def docx_to_txt_with_numbering(docx_path: str, txt_path: str) -> bool:
    """
    Converts a DOCX file to TXT with proper handling of automatic lists.

    Args:
        docx_path: Path to the input DOCX file
        txt_path: Path to the output TXT file

    Returns:
        True on success, False on error
    """
    converter = DocxToTextConverter()
    return converter.convert(docx_path, txt_path)
