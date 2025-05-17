from docx4llm import convert_docx_to_format
from docx4llm.exceptions import (
    PandocNotInstalledError,
    PandocConversionError,
)

input_file = "tests/data/output.docx"
output_format = "markdown"

try:
    converted_file_path = convert_docx_to_format(
        input_file, output_format, track_changes="all"
    )
    print(f"Successfully converted '{input_file}' to '{converted_file_path}'")
except FileNotFoundError:
    print(f"Error: Input file not found at {input_file}")
except PandocNotInstalledError as e:
    print(f"Pandoc Error: {e}")
    print(
        "Please ensure pypandoc is installed and Pandoc is in your system PATH."
    )
except PandocConversionError as e:
    print(f"Pandoc conversion failed: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
