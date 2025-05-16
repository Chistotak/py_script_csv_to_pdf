# release_notes_generator/main.py

import argparse
import logging
import sys
import os
import configparser

import csv_importer

try:
    import docx_creator
except ImportError:
    sys.exit(1)

# --- Стандартные значения ---
# ... (остаются как есть) ...
DEFAULT_COL_ISSUE_KEY = "Issue key"
DEFAULT_COL_FIX_VERSIONS_NAME = "Fix Version/s"
DEFAULT_COL_CUSTOMER_DESC = "Custom field (Description for the customer)"
DEFAULT_COL_INSTALL_INSTRUCTIONS = "Custom field (Инструкция по установке)"
DEFAULT_COL_ISSUE_TYPE = "Issue Type"
DEFAULT_CONFIG_FILE = "config.ini"
DEFAULT_FONT_NAME_DOCX = "Arial"
DEFAULT_CSV_INPUT_FILE = "input.csv"
DEFAULT_DOCX_OUTPUT_FILE = "output_releasenotes.docx"


def setup_logging():
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(name)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s',
                        handlers=[logging.StreamHandler(sys.stdout)])


def load_config(config_filepath=DEFAULT_CONFIG_FILE):
    # Получаем логгер для текущего модуля (где определена эта функция)
    # Поскольку эта функция находится в том же файле, что и main,
    # имя логгера будет __main__ если скрипт запущен напрямую.
    logger = logging.getLogger(__name__)  # <--- ДОБАВЛЕНО ЗДЕСЬ

    config = configparser.ConfigParser(allow_no_value=True, inline_comment_prefixes=('#', ';'))
    config.optionxform = str
    config_data = {
        'General': {},
        'MicroserviceVersions': {},
        'IssueTypeNames': {},
        'Columns': {}
    }
    # Устанавливаем значения по умолчанию
    config_data['General']['docx_font'] = DEFAULT_FONT_NAME_DOCX
    config_data['General']['release_title_format'] = "{{global_version}}"
    config_data['General']['use_issue_type_grouping'] = 'true'
    config_data['General']['csv_input_file'] = DEFAULT_CSV_INPUT_FILE
    config_data['General']['docx_output_file'] = DEFAULT_DOCX_OUTPUT_FILE

    config_data['Columns']['key'] = DEFAULT_COL_ISSUE_KEY
    config_data['Columns']['fix_versions'] = DEFAULT_COL_FIX_VERSIONS_NAME
    config_data['Columns']['customer_desc'] = DEFAULT_COL_CUSTOMER_DESC
    config_data['Columns']['install_instructions'] = DEFAULT_COL_INSTALL_INSTRUCTIONS
    config_data['Columns']['issue_type'] = DEFAULT_COL_ISSUE_TYPE

    if os.path.exists(config_filepath):
        try:
            config.read(config_filepath, encoding='utf-8')
            logger.info(f"Конфигурационный файл '{config_filepath}' успешно загружен.")  # Теперь logger доступен
            for section in config.sections():
                if section in config_data:
                    config_data[section].update(dict(config.items(section)))
                else:
                    config_data[section] = dict(config.items(section))
        except configparser.Error as e:
            logger.error(f"Ошибка при чтении конфигурационного файла '{config_filepath}': {e}")
    else:
        logger.warning(
            f"Конфигурационный файл '{config_filepath}' не найден. Будут использованы значения по умолчанию для путей и настроек.")

    logger.debug(f"Загруженная конфигурация: {config_data}")
    return config_data


def main():
    setup_logging()  # Сначала настраиваем логирование
    logger = logging.getLogger(__name__)  # Затем получаем логгер для функции main

    parser = argparse.ArgumentParser(description="Генератор DOCX с описанием релиза из JIRA CSV.")
    parser.add_argument("--config", default=DEFAULT_CONFIG_FILE,
                        help=f"Путь к конфигурационному файлу (по умолчанию: {DEFAULT_CONFIG_FILE})")
    parser.add_argument("--csv-file", help="Переопределить путь к входному CSV файлу (из config.ini).")
    parser.add_argument("--docx-file", help="Переопределить путь к выходному DOCX файлу (из config.ini).")
    parser.add_argument("--col-key", help=f"Переопределить имя колонки с ключом задачи")
    parser.add_argument("--col-fix-versions", help=f"Переопределить имя колонки с версиями фиксации")
    parser.add_argument("--no-issue-type-grouping", action='store_true',
                        help="Отключить группировку по типу задачи (переопределяет config.ini).")
    args = parser.parse_args()

    config_file_path = args.config

    # --- Загрузка конфигурации ---
    # Теперь load_config сама создаст/получит свой логгер
    config_data = load_config(config_file_path)

    # ... (остальная часть функции main без изменений) ...
    csv_file_path = args.csv_file if args.csv_file else config_data['General'].get('csv_input_file',
                                                                                   DEFAULT_CSV_INPUT_FILE)
    docx_file_path = args.docx_file if args.docx_file else config_data['General'].get('docx_output_file',
                                                                                      DEFAULT_DOCX_OUTPUT_FILE)

    docx_font_name = config_data['General'].get('docx_font', DEFAULT_FONT_NAME_DOCX)

    col_config = {
        'key': args.col_key if args.col_key else config_data['Columns'].get('key', DEFAULT_COL_ISSUE_KEY),
        'fix_versions_name': args.col_fix_versions if args.col_fix_versions else config_data['Columns'].get(
            'fix_versions', DEFAULT_COL_FIX_VERSIONS_NAME),
        'customer_desc': config_data['Columns'].get('customer_desc', DEFAULT_COL_CUSTOMER_DESC),
        'install_instructions': config_data['Columns'].get('install_instructions', DEFAULT_COL_INSTALL_INSTRUCTIONS),
        'issue_type': config_data['Columns'].get('issue_type', DEFAULT_COL_ISSUE_TYPE),
    }

    if args.no_issue_type_grouping:
        col_config['use_issue_type_grouping'] = False
    else:
        col_config['use_issue_type_grouping'] = config_data['General'].get('use_issue_type_grouping',
                                                                           'true').lower() == 'true'

    logger.info(f"--- Начало генерации отчета ---")
    logger.info(f"Конфигурационный файл: {os.path.abspath(config_file_path)}")
    logger.info(f"Входной CSV файл: {os.path.abspath(csv_file_path)}")
    logger.info(f"Выходной DOCX файл: {os.path.abspath(docx_file_path)}")
    logger.info(f"Конфигурация колонок: {col_config}")
    logger.info(f"Используемый шрифт для DOCX (предпочтительный): {docx_font_name}")

    raw_task_data, header_map, fix_versions_col_indices, issue_type_col_idx = \
        csv_importer.load_and_process_issues(csv_file_path, col_config)

    if raw_task_data is None:
        logger.error("Не удалось загрузить данные из CSV. DOCX не будет создан.")
        sys.exit(1)
    logger.info(f"Прочитано {len(raw_task_data)} строк задач из CSV (исключая заголовок).")

    global_version_part = csv_importer.find_global_version_title(raw_task_data, fix_versions_col_indices)

    release_title_override = config_data['General'].get('release_title_override')
    if release_title_override:
        final_release_title = release_title_override
        logger.info(f"Заголовок релиза взят из конфигурации (override): '{final_release_title}'")
    else:
        title_format_template = config_data['General'].get('release_title_format', "{{global_version}}")
        if global_version_part:
            final_release_title = title_format_template.replace("{{global_version}}", global_version_part)
        else:
            default_title_if_no_global = "Описание Релиза"
            if title_format_template and title_format_template != "{{global_version}}":
                final_release_title = title_format_template.replace("{{global_version}}", "Не указана").strip().rstrip(
                    ':').strip()
                if not final_release_title or final_release_title == config_data['General'].get('release_title_format',
                                                                                                "").replace(
                        "{{global_version}}", "").strip().rstrip(':').strip():
                    final_release_title = default_title_if_no_global
            else:
                final_release_title = default_title_if_no_global
            logger.warning(
                f"Глобальная версия не найдена, используется шаблон/заглушка для заголовка: '{final_release_title}'")
    logger.info(f"Финальный заголовок релиза для документа: '{final_release_title}'")

    logger.info("Группировка задач...")
    grouped_issues_data = csv_importer.group_issues_by_version_and_type(
        raw_task_data, header_map, col_config, fix_versions_col_indices, issue_type_col_idx, config_data
    )

    if grouped_issues_data is None:
        logger.error("Не удалось сгруппировать задачи. DOCX не будет создан.")
        sys.exit(1)

    num_ms_versions = len(grouped_issues_data)
    logger.info(f"Сгруппировано задач по {num_ms_versions} версиям микросервисов.")

    logger.info(f"Генерация DOCX документа '{docx_file_path}' для релиза '{final_release_title}'...")
    success = docx_creator.create_release_notes_docx(
        docx_file_path,
        final_release_title,
        grouped_issues_data,
        col_config['use_issue_type_grouping'],
        default_font_name=docx_font_name,
        config_data=config_data
    )

    if success:
        logger.info(f"--- Генерация отчета успешно завершена. Файл: {os.path.abspath(docx_file_path)} ---")
    else:
        logger.error("--- Возникли ошибки при создании DOCX. Файл может быть не создан или некорректен. ---")
        sys.exit(1)


if __name__ == "__main__":
    main()