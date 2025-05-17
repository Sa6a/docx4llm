"""Utilities for working with lxml and WordProcessingML XML."""

from __future__ import annotations

from typing import TYPE_CHECKING

from lxml import etree

from docx4llm.constants import NAMESPACES, WORDPROCESSINGML_NAMESPACE

if TYPE_CHECKING:
    from lxml.etree import _Element as EtreeElement


def find_element(
    parent: EtreeElement,
    xpath: str,
) -> EtreeElement | None:
    """Find a single XML element using XPath.

    Args:
        parent: The parent lxml element to search within.
        xpath: The XPath expression.

    Returns:
        The found lxml element, or None if not found.
    """
    return parent.find(xpath, namespaces=NAMESPACES)


def find_all_elements(
    parent: EtreeElement,
    xpath: str,
) -> list[EtreeElement]:
    """Find all XML elements matching an XPath expression.

    Args:
        parent: The parent lxml element or document root to search within.
        xpath: The XPath expression.

    Returns:
        A list of found lxml elements.
    """
    return parent.xpath(xpath, namespaces=NAMESPACES)  # type: ignore[no-any-return]


def get_attribute(
    element: EtreeElement | None,
    attribute_name: str,
    namespace_uri: str = WORDPROCESSINGML_NAMESPACE,
) -> str | None:
    """Get an attribute value from an XML element, supporting namespaces.

    Args:
        element: The lxml element.
        attribute_name: The local name of the attribute.
        namespace_uri: The namespace URI for the attribute.
                       Defaults to WordProcessingML namespace.

    Returns:
        The attribute value as a string, or None if not found or element is None.
    """
    if element is None:
        return None
    namespaced_attribute_name = f"{{{namespace_uri}}}{attribute_name}"
    return element.get(namespaced_attribute_name)


def create_element(
    tag_name: str,
    attributes: dict[str, str] | None = None,
    namespace_uri: str = WORDPROCESSINGML_NAMESPACE,
) -> EtreeElement:
    """Create a new namespaced XML element.

    Args:
        tag_name: The local name of the element tag.
        attributes: A dictionary of attributes to set on the element.
                    Keys are local attribute names, values are attribute values.
                    These attributes will also be namespaced with namespace_uri.
        namespace_uri: The namespace URI for the element and its attributes.
                       Defaults to WordProcessingML namespace.

    Returns:
        The created lxml element.
    """
    full_tag_name = f"{{{namespace_uri}}}{tag_name}"
    element = etree.Element(full_tag_name, nsmap=NAMESPACES)

    if attributes:
        for attr_name, attr_value in attributes.items():
            full_attr_name = f"{{{namespace_uri}}}{attr_name}"
            element.set(full_attr_name, attr_value)
    return element
