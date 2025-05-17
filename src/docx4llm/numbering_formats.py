"""Defines number formatting strategies and a service to manage them."""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Final


class INumberFormatter(ABC):
    """Interface for number formatting strategies."""

    @abstractmethod
    def format(self, number: int) -> str:
        """Format the given integer into a string representation.

        Args:
            number: The integer to format.

        Returns:
            The string representation of the number.
        """


class _RomanNumeralHelper:
    """Helper class for Roman numeral conversion logic."""

    _ROMAN_MAP: Final[dict[int, str]] = {
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

    @staticmethod
    def to_roman(number: int) -> str:
        """Convert an integer to its Roman numeral representation.

        Args:
            number: The integer to convert. Must be positive.

        Returns:
            The Roman numeral string.

        Raises:
            ValueError: If the number is not positive or too large for typical
                        Roman numeral representation (though this implementation
                        handles large numbers by repetition).
        """
        if number <= 0:
            raise ValueError(
                "Roman numeral conversion requires a positive integer."
            )

        result_parts = []
        for value, numeral in _RomanNumeralHelper._ROMAN_MAP.items():
            while number >= value:
                result_parts.append(numeral)
                number -= value
        return "".join(result_parts)


class DecimalFormatter(INumberFormatter):
    """Formats numbers as decimal strings."""

    def format(self, number: int) -> str:
        """Format number as decimal. E.g., 1, 2, 3."""
        return str(number)


class UpperRomanFormatter(INumberFormatter):
    """Formats numbers as uppercase Roman numerals."""

    def format(self, number: int) -> str:
        """Format number as uppercase Roman. E.g., I, II, III."""
        return _RomanNumeralHelper.to_roman(number)


class LowerRomanFormatter(INumberFormatter):
    """Formats numbers as lowercase Roman numerals."""

    def format(self, number: int) -> str:
        """Format number as lowercase Roman. E.g., i, ii, iii."""
        return _RomanNumeralHelper.to_roman(number).lower()


class UpperLetterFormatter(INumberFormatter):
    """Formats numbers as uppercase letters."""

    def format(self, number: int) -> str:
        """Format number as uppercase letter. E.g., A, B, C.

        Assumes 1-based indexing for letters (1=A, 2=B, ...).
        """
        if number < 1 or number > 26:

            return str(number)
        return chr(ord("A") + number - 1)


class LowerLetterFormatter(INumberFormatter):
    """Formats numbers as lowercase letters."""

    def format(self, number: int) -> str:
        """Format number as lowercase letter. E.g., a, b, c.

        Assumes 1-based indexing for letters (1=a, 2=b, ...).
        """
        if number < 1 or number > 26:
            return str(number)
        return chr(ord("a") + number - 1)


class NumberingFormatterService:
    """Provides access to various number formatters."""

    def __init__(self) -> None:
        """Initialize the service with available formatters."""
        self._formatters: dict[str, INumberFormatter] = {
            "decimal": DecimalFormatter(),
            "upperRoman": UpperRomanFormatter(),
            "lowerRoman": LowerRomanFormatter(),
            "upperLetter": UpperLetterFormatter(),
            "lowerLetter": LowerLetterFormatter(),
        }
        self._default_formatter: INumberFormatter = DecimalFormatter()

    def get_formatter(self, format_type: str) -> INumberFormatter:
        """Retrieve a formatter for the specified type.

        Args:
            format_type: The type of formatter desired (e.g., "decimal").

        Returns:
            An instance of INumberFormatter. Returns a default (decimal)
            formatter if the requested type is not found.
        """
        return self._formatters.get(format_type, self._default_formatter)
