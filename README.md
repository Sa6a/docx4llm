## Функция `add_numbering_to_docx`

**Назначение:**
Функция add_numbering_to_docx преобразует нумерованные списки в документе Word в текстовый формат, что делает их более доступными для обработки с помощью языковых моделей.

**Использование:**

```python
import main

input_path = "путь/к/вашему/входному_документу.docx"
output_path = "путь/к/вашему/выходному_документу.docx"

success = main.add_numbering_to_docx(input_path, output_path)

if success:
    print(f"Нумерация успешно добавлена в {output_path}")
else:
    print(f"Не удалось добавить нумерацию в {input_path}")

```

## Функция `convert_docx_to_format`

**Назначение:**
Функция convert_docx_to_format позволяет конвертировать документы Word в различные форматы, такие как Markdown, HTML, PDF и другие.

**Зависимости:**
Для работы этой функции требуется установленная библиотека `pypandoc`. Также необходимо установить Pandoc: https://pandoc.org/installing.html

**Использование:**

```python
import main

input_path = "путь/к/вашему/документу.docx"
output_format_type = "pdf"  # Например, 'html', 'md', 'odt' и т.д.

# Пример с принятием всех изменений
# output_file = main.convert_docx_to_format(input_path, output_format_type, track_changes="accept")

# Пример с отклонением всех изменений
# output_file = main.convert_docx_to_format(input_path, output_format_type, track_changes="reject")

# Пример с сохранением информации об изменениях (по умолчанию)
output_file = main.convert_docx_to_format(input_path, output_format_type, track_changes="all")


if output_file:
    print(f"Файл успешно сконвертирован в {output_format_type}: {output_file}")
else:
    print(f"Не удалось сконвертировать файл {input_path}")
```
