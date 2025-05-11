import os
import shutil
import zipfile
from typing import Any

from lxml import etree


class NumberingFormat:
    @staticmethod
    def format_number(number: int, format_type: str) -> str:
        if format_type == "decimal":
            return str(number)
        if format_type == "upperRoman":
            return NumberingFormat._to_roman(number)
        if format_type == "lowerRoman":
            return NumberingFormat._to_roman(number).lower()
        if format_type == "upperLetter":
            return chr(ord("A") + number - 1)
        if format_type == "lowerLetter":
            return chr(ord("a") + number - 1)
        return str(number)

    @staticmethod
    def _to_roman(number: int) -> str:
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
    def __init__(self, format_type: str, text_template: str, start_value: int):
        self.format_type = format_type
        self.text_template = text_template
        self.start_value = start_value
        self.current_value = start_value

    def reset(self) -> None:
        self.current_value = self.start_value

    def increment(self) -> None:
        self.current_value += 1

    def format_current_value(self) -> str:
        return NumberingFormat.format_number(
            self.current_value,
            self.format_type,
        )


class NumberingDefinition:
    def __init__(self, abstract_num_id: str):
        self.abstract_num_id = abstract_num_id
        self.levels: dict[str, NumberingLevel] = {}

    def add_level(self, level_id: str, level: NumberingLevel) -> None:
        self.levels[level_id] = level

    def reset_levels_below(self, current_level_id: str) -> None:
        current_level_int = int(current_level_id)
        for level_id, level in self.levels.items():
            if int(level_id) > current_level_int:
                level.reset()

    def get_formatted_number(self, level_id: str) -> str:
        if level_id not in self.levels:
            return ""

        level = self.levels[level_id]
        text = level.text_template

        level_ids = sorted(self.levels.keys(), key=lambda x: int(x))
        for sub_level_id in level_ids:
            sub_level = self.levels[sub_level_id]
            placeholder = f"%{int(sub_level_id) + 1}"
            if placeholder in text:
                text = text.replace(
                    placeholder,
                    sub_level.format_current_value(),
                )

        return text


class XmlHelper:
    WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_MAP = {"w": WORD_NS}

    @staticmethod
    def find_element(parent, xpath: str) -> etree.Element | None:
        return parent.find(xpath, namespaces=XmlHelper.NS_MAP)

    @staticmethod
    def find_all_elements(parent, xpath: str) -> list[etree.Element]:
        return parent.xpath(xpath, namespaces=XmlHelper.NS_MAP)

    @staticmethod
    def get_attribute(element, attr_name: str) -> str | None:
        if element is None:
            return None

        full_attr_name = f"{{{XmlHelper.WORD_NS}}}{attr_name}"
        return element.get(full_attr_name)

    @staticmethod
    def parse_numbering_info(
        numPr_element,
    ) -> tuple[str | None, str | None]:
        if numPr_element is None:
            return None, None

        ilvl_element = XmlHelper.find_element(numPr_element, ".//w:ilvl")
        num_id_element = XmlHelper.find_element(numPr_element, ".//w:numId")

        ilvl = XmlHelper.get_attribute(ilvl_element, "val")
        num_id = XmlHelper.get_attribute(num_id_element, "val")

        return ilvl, num_id

    @staticmethod
    def has_element(parent, xpath: str) -> bool:
        return len(XmlHelper.find_all_elements(parent, xpath)) > 0

    @staticmethod
    def element_exists(element, xpath: str) -> bool:
        return XmlHelper.find_element(element, xpath) is not None

    @staticmethod
    def create_element(
        tag_name: str,
        attributes: dict[str, str] | None = None,
    ) -> etree.Element:
        full_tag_name = f"{{{XmlHelper.WORD_NS}}}{tag_name}"
        element = etree.Element(full_tag_name)

        if attributes:
            for attr_name, attr_value in attributes.items():
                full_attr_name = f"{{{XmlHelper.WORD_NS}}}{attr_name}"
                element.set(full_attr_name, attr_value)

        return element


class NumberingParser:
    def __init__(self):
        self.abstract_numbering_data: dict[str, dict[str, dict[str, Any]]] = {}
        self.numbering_definitions: dict[str, NumberingDefinition] = {}

    def parse_numbering_xml(self, numbering_xml_content: bytes) -> None:
        numbering_root = etree.fromstring(numbering_xml_content)

        self._parse_abstract_numbering(numbering_root)
        self._parse_numbering(numbering_root)

    def _parse_abstract_numbering(self, numbering_root: etree.Element) -> None:
        for abstract_num in XmlHelper.find_all_elements(
            numbering_root,
            "//w:abstractNum",
        ):
            abstract_num_id = XmlHelper.get_attribute(
                abstract_num,
                "abstractNumId",
            )
            if not abstract_num_id:
                continue

            self.abstract_numbering_data[abstract_num_id] = {}

            for lvl in XmlHelper.find_all_elements(abstract_num, ".//w:lvl"):
                ilvl = XmlHelper.get_attribute(lvl, "ilvl")
                if not ilvl:
                    continue

                num_fmt_element = XmlHelper.find_element(lvl, ".//w:numFmt")
                num_fmt = (
                    XmlHelper.get_attribute(num_fmt_element, "val") or "decimal"
                )

                lvl_text_element = XmlHelper.find_element(lvl, ".//w:lvlText")
                lvl_text = (
                    XmlHelper.get_attribute(lvl_text_element, "val") or "%1."
                )

                start_element = XmlHelper.find_element(lvl, ".//w:start")
                start_str = XmlHelper.get_attribute(start_element, "val")
                start = int(start_str) if start_str else 1

                self.abstract_numbering_data[abstract_num_id][ilvl] = {
                    "format": num_fmt,
                    "text": lvl_text,
                    "start": start,
                }

    def _parse_numbering(self, numbering_root: etree.Element) -> None:
        for num in XmlHelper.find_all_elements(numbering_root, "//w:num"):
            num_id = XmlHelper.get_attribute(num, "numId")
            if not num_id:
                continue

            abstract_num_id_element = XmlHelper.find_element(
                num,
                ".//w:abstractNumId",
            )
            abstract_num_id = XmlHelper.get_attribute(
                abstract_num_id_element,
                "val",
            )
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
                    lvl_data["format"],
                    lvl_data["text"],
                    lvl_data["start"],
                )
                num_def.add_level(lvl_id, level)

            for lvl_override in XmlHelper.find_all_elements(
                num,
                ".//w:lvlOverride",
            ):
                ilvl = XmlHelper.get_attribute(lvl_override, "ilvl")
                if not ilvl or ilvl not in num_def.levels:
                    continue

                start_override = XmlHelper.find_element(
                    lvl_override,
                    ".//w:startOverride",
                )
                if start_override:
                    new_start_str = XmlHelper.get_attribute(
                        start_override,
                        "val",
                    )
                    if new_start_str:
                        new_start = int(new_start_str)
                        num_def.levels[ilvl].start_value = new_start
                        num_def.levels[ilvl].current_value = new_start

            self.numbering_definitions[num_id] = num_def


class ParagraphFormatter:
    def __init__(self, numbering_definitions: dict[str, NumberingDefinition]):
        self.numbering_definitions = numbering_definitions
        self.last_active_levels = {}

    def format_paragraph(self, paragraph: etree.Element) -> str:
        pPr = XmlHelper.find_element(paragraph, ".//w:pPr")
        if pPr is None:
            return ""

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

        return num_prefix

    def _has_valid_numbering(
        self,
        ilvl: str | None,
        num_id: str | None,
    ) -> bool:
        if not ilvl or not num_id or num_id not in self.numbering_definitions:
            return False

        if ilvl not in self.numbering_definitions[num_id].levels:
            return False

        return True


class DocumentCleaner:
    @staticmethod
    def clean_document(document_root: etree.Element) -> None:
        xpath = "//w:pPrChange//w:numPr | //w:rPrChange//w:numPr | //w:p[@w:rsidDel]//w:numPr"
        for numPr in XmlHelper.find_all_elements(document_root, xpath):
            parent = numPr.getparent()
            if parent is not None:
                parent.remove(numPr)


class DocxNumberingProcessor:
    def __init__(self):
        self.numbering_parser = NumberingParser()

    def process(self, input_docx_path: str, output_docx_path: str) -> bool:
        temp_dir = f"{input_docx_path}_temp"

        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            with zipfile.ZipFile(input_docx_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)

            self._process_files(temp_dir)

            self._create_docx(temp_dir, output_docx_path)

            shutil.rmtree(temp_dir)

            return True
        except Exception as e:
            print(f"Ошибка при обработке файла: {e}")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            return False

    def _process_files(self, temp_dir: str) -> None:
        numbering_path = os.path.join(temp_dir, "word", "numbering.xml")
        if os.path.exists(numbering_path):
            with open(numbering_path, "rb") as f:
                self.numbering_parser.parse_numbering_xml(f.read())

        document_path = os.path.join(temp_dir, "word", "document.xml")
        if os.path.exists(document_path):
            with open(document_path, "rb") as f:
                document_content = f.read()

            modified_document = self._process_document(document_content)

            with open(document_path, "wb") as f:
                f.write(modified_document)

    def _process_document(self, document_content: bytes) -> bytes:
        document_root = etree.fromstring(document_content)

        DocumentCleaner.clean_document(document_root)

        paragraph_formatter = ParagraphFormatter(
            self.numbering_parser.numbering_definitions,
        )

        for paragraph in XmlHelper.find_all_elements(document_root, "//w:p"):
            num_prefix = paragraph_formatter.format_paragraph(paragraph)

            if num_prefix:
                self._add_numbering_to_paragraph(paragraph, num_prefix)

            self._remove_numPr_tags(paragraph)

        return etree.tostring(
            document_root,
            encoding="utf-8",
            xml_declaration=True,
        )

    def _add_numbering_to_paragraph(
        self,
        paragraph: etree.Element,
        num_prefix: str,
    ) -> None:
        first_t = XmlHelper.find_element(paragraph, ".//w:t")

        if first_t is not None:
            current_text = first_t.text or ""
            first_t.text = f"{num_prefix} {current_text}"
        else:
            r_element = XmlHelper.create_element("r")
            t_element = XmlHelper.create_element("t")
            t_element.text = f"{num_prefix}"

            r_element.append(t_element)
            paragraph.append(r_element)

    def _remove_numPr_tags(self, paragraph: etree.Element) -> None:
        for numPr in XmlHelper.find_all_elements(paragraph, ".//w:numPr"):
            parent = numPr.getparent()
            if parent is not None:
                parent.remove(numPr)

    def _create_docx(self, source_dir: str, output_path: str) -> None:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, source_dir)
                    zipf.write(file_path, arcname)


def add_numbering_to_docx(input_docx_path: str, output_docx_path: str) -> bool:
    processor = DocxNumberingProcessor()
    return processor.process(input_docx_path, output_docx_path)


def convert_docx_to_format(
    input_docx_path: str,
    output_format: str,
    track_changes: str = "all",
) -> str:

    try:
        import pypandoc

        output_file_path = (
            os.path.splitext(input_docx_path)[0] + f".{output_format}"
        )

        extra_args = []

        if track_changes not in ["accept", "reject", "all"]:
            print(
                f"Предупреждение: некорректное значение track_changes '{track_changes}'. Используется 'all'.",
            )
            track_changes = "all"

        extra_args.append(f"--track-changes={track_changes}")

        pypandoc.convert_file(
            input_docx_path,
            output_format,
            outputfile=output_file_path,
            extra_args=extra_args,
        )

        print(
            f"Файл успешно сконвертирован в формат {output_format}: {output_file_path}",
        )
        return output_file_path

    except ImportError:
        print(
            "Ошибка: библиотека pypandoc не установлена. Установите её командой: pip install pypandoc",
        )
        return ""
    except Exception as e:
        print(f"Ошибка при конвертации файла: {e}")
        return ""
