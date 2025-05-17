"""Command Line Interface for DOCX Numbering Processor."""

import logging
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import cast

try:
    from docx4llm import (
        DocxProcessingError,
        PandocConversionError,
        PandocNotInstalledError,
        add_numbering_to_docx,
        convert_docx_to_format,
    )
    from docx4llm.pandoc_utils import TrackChangesOption
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from docx4llm import (
        DocxProcessingError,
        PandocConversionError,
        PandocNotInstalledError,
        add_numbering_to_docx,
        convert_docx_to_format,
    )
    from docx4llm.pandoc_utils import TrackChangesOption


ERROR_LOG_FILENAME = "DocxNumConvert_error.log"
_error_file_logger: logging.Logger | None = None


def _get_error_file_logger() -> logging.Logger:
    """Initialize and return the dedicated error file logger."""
    global _error_file_logger
    if _error_file_logger is None:
        _error_file_logger = logging.getLogger("DocxNumConvertErrorFile")
        _error_file_logger.setLevel(logging.ERROR)
        _error_file_logger.propagate = False

        if not any(
            isinstance(h, logging.FileHandler)
            and Path(getattr(h, 'baseFilename', ''))
            == Path.cwd() / ERROR_LOG_FILENAME
            for h in _error_file_logger.handlers
        ):
            try:
                file_handler = logging.FileHandler(
                    ERROR_LOG_FILENAME,
                    mode="a",
                    encoding="utf-8",
                )
                file_formatter = logging.Formatter(
                    "[%(asctime)s] %(message)s",
                    datefmt="%Y-%m-%d %H:%M:%S",
                )
                file_handler.setFormatter(file_formatter)
                _error_file_logger.addHandler(file_handler)
            except IOError as e:
                print(
                    f"КРИТИЧЕСКАЯ ОШИБКА: Не удалось инициализировать лог-файл "
                    f"'{ERROR_LOG_FILENAME}': {e!s}",
                    file=sys.stderr,
                )
    return _error_file_logger


def log_error_and_exit(
    message: str,
    original_err: Exception | None = None,
) -> None:
    """Log an error message to console and file, then exit."""
    error_logger = _get_error_file_logger()

    log_message_to_file_parts = [message]
    if original_err:
        log_message_to_file_parts.append(
            f"\n  Подробности ошибки: {original_err!s}",
        )
    log_message_to_file_parts.append("\n")

    full_log_message = "".join(log_message_to_file_parts)

    print("-----------------------------------------", file=sys.stderr)
    print(f"ОШИБКА: {message}", file=sys.stderr)
    if original_err:
        print(f"  Подробности: {original_err!s}", file=sys.stderr)

    try:
        error_logger.error(full_log_message.strip())
        print(
            f"  Подробности этой ошибки сохранены в файле: {ERROR_LOG_FILENAME}",
            file=sys.stderr,
        )
    except Exception as e_log:
        print(
            f"  КРИТИЧЕСКАЯ ПОД-ОШИБКА: Не удалось записать детали в лог-файл "
            f"'{ERROR_LOG_FILENAME}': {e_log!s}",
            file=sys.stderr,
        )
    print("-----------------------------------------", file=sys.stderr)
    sys.exit(1)


def get_user_input(prompt: str) -> str:
    """Get trimmed user input from console."""
    return input(prompt).strip()


def ask_yes_no(prompt: str) -> bool:
    """Ask a yes/no question and return boolean response."""
    while True:
        answer = get_user_input(f"{prompt} (да/нет): ").lower()
        if answer in {"да", "д", "yes", "y"}:
            return True
        if answer in {"нет", "н", "no", "n"}:
            return False
        print("Пожалуйста, ответьте 'да' или 'нет'.")


def check_pandoc_installed() -> bool:
    """Check if Pandoc executable is found in PATH."""
    return shutil.which("pandoc") is not None


def main_cli() -> None:
    """Main function for the Command Line Interface."""
    print("--- Обработчик нумерации DOCX ---")

    input_docx_path_str: str
    while True:
        input_docx_path_str = get_user_input(
            "Введите полный путь к DOCX файлу для обработки: ",
        )
        input_docx_path = Path(input_docx_path_str)
        if not input_docx_path_str:
            print("Путь к файлу не может быть пустым.")
            continue

        if input_docx_path.is_file():
            if input_docx_path.suffix.lower() == ".docx":
                break
            print("Файл должен иметь расширение .docx")
        elif not input_docx_path.exists():
            print(f"Файл не найден: {input_docx_path_str}")
        else:
            log_error_and_exit(
                f"Ошибка при проверке файла '{input_docx_path_str}'. "
                "Указанный путь не является файлом.",
                None,
            )
        print("Пожалуйста, попробуйте снова.")

    input_path = Path(input_docx_path_str)
    dir_name = input_path.parent
    name_without_ext = input_path.stem
    ext = input_path.suffix

    output_docx_processed_path = dir_name / f"{name_without_ext}_numbered{ext}"

    print(f"Файл будет обработан и сохранен как: {output_docx_processed_path}")
    print(f"Начинаю обработку файла: {input_path}...")

    try:
        success = add_numbering_to_docx(input_path, output_docx_processed_path)
        if not success:
            log_error_and_exit(
                f"Не удалось обработать файл '{input_path}'. "
                "Функция add_numbering_to_docx вернула false без явной ошибки.",
                None,
            )
    except FileNotFoundError as err:
        log_error_and_exit(
            f"Файл не найден при попытке обработки: {err.filename}", err
        )
    except DocxProcessingError as err:
        log_error_and_exit(
            f"Ошибка при обработке DOCX файла '{input_path}'", err
        )
    except Exception as err:
        log_error_and_exit(
            f"Неожиданная ошибка при обработке DOCX файла '{input_path}'",
            err,
        )

    print(
        f"Файл '{input_path}' успешно обработан и сохранен как "
        f"'{output_docx_processed_path}'",
    )

    if not ask_yes_no(
        "Хотите сконвертировать обработанный DOCX файл в другой формат?",
    ):
        print("Завершение работы.")
        return

    if not check_pandoc_installed():
        err_msg = (
            "Pandoc не найден. Для конвертации файлов установите Pandoc "
            "(https://pandoc.org/installing.html)."
        )
        log_error_and_exit(err_msg, None)

    print(
        "\nДоступные форматы для конвертации (некоторые могут требовать "
        "доп. ПО, например, LaTeX для PDF):",
    )
    print(" - markdown (рекомендуется)")
    print(" - gfm (GitHub Flavored Markdown)")
    print(" - html")
    print(" - pdf")
    print(" - commonmark")
    print("Или введите любой другой формат, поддерживаемый Pandoc.")
    print(
        "ПРЕДУПРЕЖДЕНИЕ: Pandoc не умеет работать с некоторыми таблицами и "
        "будет удалять их содержимое.",
    )

    output_format = ""
    while not output_format:
        output_format = get_user_input("Введите желаемый формат вывода: ")
        if not output_format:
            print("Формат не может быть пустым.")

    default_track_changes: TrackChangesOption = "all"
    print("\nРежим отслеживания изменений при конвертации:")
    print(f" - {default_track_changes} (сохранить все изменения и комментарии)")
    print(" - accept (принять все изменения)")
    print(" - reject (отклонить все изменения)")

    user_track_changes_str = get_user_input(
        f"Введите режим отслеживания изменений (или Enter для '{default_track_changes}'): ",
    )

    track_changes_to_use: TrackChangesOption = default_track_changes
    if user_track_changes_str:
        valid_literal_options: list[TrackChangesOption] = [
            "all",
            "accept",
            "reject",
        ]
        if user_track_changes_str.lower() in valid_literal_options:
            track_changes_to_use = cast(
                TrackChangesOption,
                user_track_changes_str.lower(),
            )
        else:
            print(
                f"Предупреждение: Введенный режим '{user_track_changes_str}' "
                f"не распознан. Будет использован режим '{default_track_changes}'.",
            )

    print(
        f"\nНачинаю конвертацию файла '{output_docx_processed_path}' "
        f"в формат '{output_format}'...",
    )

    try:
        converted_file_path = convert_docx_to_format(
            output_docx_processed_path,
            output_format,
            track_changes=track_changes_to_use,
        )
        print(
            f"Файл успешно сконвертирован и сохранен как: {converted_file_path}"
        )
    except PandocNotInstalledError as err:
        log_error_and_exit(str(err), err)
    except PandocConversionError as err:
        err_msg_for_log = (
            f"Ошибка при конвертации файла '{output_docx_processed_path}' "
            f"в формат '{output_format}'.\n"
            "Возможная команда Pandoc: "
            f'pandoc "{output_docx_processed_path}" -f docx -t {output_format} '
            f'-o ... --track-changes={track_changes_to_use}'
        )
        log_error_and_exit(err_msg_for_log, err)
    except FileNotFoundError as err:
        log_error_and_exit(
            f"Файл для конвертации не найден: {err.filename}", err
        )
    except Exception as err:
        log_error_and_exit(
            f"Неожиданная ошибка при конвертации файла '{output_docx_processed_path}'",
            err,
        )

    print("Завершение работы.")


if __name__ == "__main__":
    _get_error_file_logger()
    main_cli()
