# release_notes_generator/main.py

import argparse  # Оставим импорт, если захотите вернуть
import logging
import sys
import os

import csv_importer

try:
    import docx_creator
except ImportError:
    sys.exit(1)

# --- Стандартные имена колонок (если не переопределены) ---
DEFAULT_COL_ISSUE_KEY = "Issue key"
DEFAULT_COL_FIX_VERSIONS_NAME = "Fix Version/s"
DEFAULT_COL_CUSTOMER_DESC = "Custom field (Description for the customer)"
DEFAULT_COL_INSTALL_INSTRUCTIONS = "Custom field (Инструкция по установке)"
DEFAULT_COL_ISSUE_TYPE = "Issue Type"
DEFAULT_OUTPUT_FILENAME = "release_notes.docx"  # Это уже не будет использоваться при хардкоде
DEFAULT_FONT_NAME_DOCX = "Arial"


def setup_logging():
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        handlers=[logging.StreamHandler(sys.stdout)])


def main():
    setup_logging()
    logger = logging.getLogger(__name__)

    # --- ЗАХАРДКОЖЕННЫЕ ПАРАМЕТРЫ ---
    logger.warning("Используются захардкоженные параметры для отладки!")

    # Укажите абсолютные или относительные пути.
    # Относительные пути будут отсчитываться от директории, где лежит main.py
    # Пример для Windows: csv_file_path = r"C:\Users\haranski\PycharmProjects\PythonProject\release.csv"
    # Пример для Linux/macOS: csv_file_path = "/home/user/projects/release_notes_generator/release.csv"

    csv_file_path = "release.csv"  # <--- ЗАМЕНИТЕ НА ВАШ ПУТЬ К CSV, если он не лежит рядом с main.py
    docx_file_path = "release_notes_hardcoded.docx"  # <--- ИМЯ ВЫХОДНОГО ФАЙЛА (будет создан рядом с main.py)
    docx_font_name = "Arial"  # Или "DejaVu Sans", "Calibri" и т.д.

    # Настройки колонок (измените, если ваши заголовки в CSV другие)
    col_key_name = "Issue key"
    col_fix_versions_name = "Fix Version/s"
    col_customer_desc_name = "Custom field (Description for the customer)"
    col_install_instructions_name = "Custom field (Инструкция по установке)"
    col_issue_type_name = "Issue Type"

    # Включить или отключить группировку по типу задачи
    enable_issue_type_grouping = True

    col_config = {
        'key': col_key_name,
        'fix_versions_name': col_fix_versions_name,
        'customer_desc': col_customer_desc_name,
        'install_instructions': col_install_instructions_name,
        'issue_type': col_issue_type_name,
        'use_issue_type_grouping': enable_issue_type_grouping
    }
    # --- КОНЕЦ БЛОКА ЗАХАРДКОЖЕННЫХ ПАРАМЕТРОВ ---

    # --- Блок argparse (закомментирован) ---
    """
    parser = argparse.ArgumentParser(description="Генератор DOCX с описанием релиза из JIRA CSV.")
    # ... (все аргументы argparse) ...
    args = parser.parse_args()

    csv_file_path = args.csv_file
    docx_file_path = args.docx_file
    docx_font_name = args.docx_font

    col_config = {
        'key': args.col_key,
        'fix_versions_name': args.col_fix_versions,
        'customer_desc': args.col_customer_desc,
        'install_instructions': args.col_install_instructions,
        'issue_type': args.col_issue_type,
        'use_issue_type_grouping': not args.no_issue_type_grouping
    }
    """
    # --- Конец блока argparse ---

    logger.info(f"--- Начало генерации отчета ---")
    logger.info(f"Входной CSV файл: {os.path.abspath(csv_file_path)}")  # os.path.abspath покажет полный путь
    logger.info(f"Выходной DOCX файл: {os.path.abspath(docx_file_path)}")
    logger.info(f"Конфигурация колонок: {col_config}")
    logger.info(f"Используемый шрифт для DOCX (предпочтительный): {docx_font_name}")

    raw_task_data, header_map, fix_versions_col_indices, issue_type_col_idx = \
        csv_importer.load_and_process_issues(csv_file_path, col_config)

    if raw_task_data is None:
        logger.error("Не удалось загрузить данные из CSV. DOCX не будет создан.")
        sys.exit(1)
    logger.info(f"Прочитано {len(raw_task_data)} строк задач из CSV (исключая заголовок).")

    docx_main_title_base = csv_importer.find_global_version_title(raw_task_data, fix_versions_col_indices)

    logger.info("Группировка задач...")
    grouped_issues_data = csv_importer.group_issues_by_version_and_type(
        raw_task_data, header_map, col_config, fix_versions_col_indices, issue_type_col_idx
    )

    if grouped_issues_data is None:
        logger.error("Не удалось сгруппировать задачи. DOCX не будет создан.")
        sys.exit(1)

    num_ms_versions = len(grouped_issues_data)
    logger.info(f"Сгруппировано задач по {num_ms_versions} версиям микросервисов.")
    if not grouped_issues_data:
        logger.warning(
            "Не найдено задач, подходящих под критерии группировки. DOCX может содержать только заголовок и дату.")
    else:
        for ms_ver, data in grouped_issues_data.items():
            if col_config['use_issue_type_grouping'] and isinstance(data, dict):  # Проверка, что data - это словарь
                for issue_type, tasks in data.items():
                    logger.debug(f"  Версия {ms_ver}, Тип '{issue_type}': {len(tasks)} задач(и)")
            elif isinstance(data, list):  # Если use_issue_type_grouping=False, data будет списком
                logger.debug(f"  Версия {ms_ver}: {len(data)} задач(и)")
            else:  # На всякий случай
                logger.debug(f"  Версия {ms_ver}: структура данных не определена для логирования ({type(data)})")

    logger.info(f"Генерация DOCX документа '{docx_file_path}' для релиза '{docx_main_title_base}'...")
    success = docx_creator.create_release_notes_docx(
        docx_file_path,
        docx_main_title_base,
        grouped_issues_data,
        col_config['use_issue_type_grouping'],
        default_font_name=docx_font_name
    )

    if success:
        logger.info(f"--- Генерация отчета успешно завершена. Файл: {os.path.abspath(docx_file_path)} ---")
    else:
        logger.error("--- Возникли ошибки при создании DOCX. Файл может быть не создан или некорректен. ---")
        sys.exit(1)


if __name__ == "__main__":
    main()