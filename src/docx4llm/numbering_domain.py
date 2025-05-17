"""models for DOCX numbering."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx4llm.numbering_formats import INumberFormatter


class NumberingLevel:
    """Represents a single level within a numbering definition."""

    def __init__(
        self,
        level_id: str,
        formatter: INumberFormatter,
        text_template: str,
        start_value: int,
    ) -> None:
        """Initialize a numbering level.

        Args:
            level_id: The identifier for this level (e.g., "0", "1").
            formatter: The formatter to use for this level's numbers.
            text_template: The template string for the numbering text,
                           e.g., "%1.%2.".
            start_value: The initial value for this numbering level.
        """
        self.level_id = level_id
        self.formatter = formatter
        self.text_template = text_template
        self.start_value = start_value
        self.current_value: int = start_value

    def reset(self) -> None:
        """Reset the current value of this level to its start value."""
        self.current_value = self.start_value

    def increment(self) -> None:
        """Increment the current value of this level by one."""
        self.current_value += 1

    def format_current_value(self) -> str:
        """Format the current value of this level using its formatter.

        Returns:
            The formatted string representation of the current value.
        """
        return self.formatter.format(self.current_value)


class NumberingDefinition:
    """Represents an abstract numbering definition and its levels."""

    def __init__(self, abstract_num_id: str) -> None:
        """Initialize a numbering definition.

        Args:
            abstract_num_id: The ID of the abstract numbering definition
                             this definition is based on.
        """
        self.abstract_num_id = abstract_num_id
        self.levels: dict[str, NumberingLevel] = {}

    def add_level(self, level: NumberingLevel) -> None:
        """Add a numbering level to this definition.

        Args:
            level: The NumberingLevel instance to add.
        """
        self.levels[level.level_id] = level

    def reset_levels_below(self, current_level_id_str: str) -> None:
        """Reset all numbering levels numerically greater than the current one.

        Args:
            current_level_id_str: The ID of the current level (e.g., "0").
                                  Levels with IDs like "1", "2" will be reset.
        """
        try:
            current_level_int = int(current_level_id_str)
        except ValueError:
            return

        for level_id, level in self.levels.items():
            try:
                level_id_int = int(level_id)
                if level_id_int > current_level_int:
                    level.reset()
            except ValueError:
                continue

    def get_formatted_number(self, target_level_id: str) -> str:
        """Generate the formatted number string for a specific level.

        This involves substituting placeholders in the target level's text
        template with the formatted current values of relevant parent/self levels.

        Args:
            target_level_id: The ID of the level for which to format the number.

        Returns:
            The fully formatted number string (e.g., "1.a.i.").
            Returns an empty string if the target_level_id is not found.
        """
        target_level = self.levels.get(target_level_id)
        if not target_level:
            return ""

        text = target_level.text_template

        sorted_level_ids = sorted(self.levels.keys(), key=int)

        for sub_level_id_str in sorted_level_ids:
            sub_level = self.levels[sub_level_id_str]
            placeholder = f"%{int(sub_level_id_str) + 1}"
            if placeholder in text:
                text = text.replace(
                    placeholder,
                    sub_level.format_current_value(),
                )
        return text
