from docx4llm import add_numbering_to_docx, DocxProcessingError

input_file = "tests/data/test_docx.docx"
output_file = "tests/data/output.docx"

try:
    success = add_numbering_to_docx(input_file, output_file)
    if success:
        print(f"Successfully processed '{input_file}' to '{output_file}'")
    else:
        print(f"Failed to process '{input_file}'")
except FileNotFoundError:
    print(f"Error: Input file not found at {input_file}")
except DocxProcessingError as e:
    print(f"An error occurred during DOCX processing: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
