# release_notes_generator/main.py

import argparse
import logging
import sys
import os
import configparser
import re  # Для extract_microservice_info_for_summary

import csv_importer

try:
    import docx_creator
except ImportError:
    sys.exit(1)

# --- Стандартные значения ---
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
    logger = logging.getLogger(__name__)
    config = configparser.ConfigParser(allow_no_value=True, inline_comment_prefixes=('#', ';'))
    config.optionxform = str  # Сохраняем регистр ключей опций

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
            logger.info(f"Конфигурационный файл '{config_filepath}' успешно загружен.")
            for section_name_from_file in config.sections():
                # Ищем соответствие секции из файла в нашем config_data
                # (с учетом возможного разного регистра, хотя optionxform=str должен помочь)
                matched_section_key = None
                for key_in_config_data in config_data.keys():
                    if key_in_config_data.lower() == section_name_from_file.lower():
                        matched_section_key = key_in_config_data
                        break

                if matched_section_key:
                    config_data[matched_section_key].update(dict(config.items(section_name_from_file)))
                else:
                    # Если секции нет в нашем шаблоне, добавляем ее как есть
                    config_data[section_name_from_file] = dict(config.items(section_name_from_file))
        except configparser.Error as e:
            logger.error(f"Ошибка при чтении конфигурационного файла '{config_filepath}': {e}")
    else:
        logger.warning(
            f"Конфигурационный файл '{config_filepath}' не найден. Будут использованы значения по умолчанию.")

    logger.debug(f"Загруженная конфигурация: {config_data}")
    return config_data


def extract_microservice_info_for_summary(grouped_data_keys, config_data):
    logger = logging.getLogger(__name__)
    microservices_summary = []
    version_pattern = re.compile(r"^([A-Z]{2})(\d+(\.\d+){1,2})$")
    seen_summary_entries = set()

    for original_ms_key in sorted(list(grouped_data_keys)):
        service_name_for_table = original_ms_key
        version_number_part = ""

        match = version_pattern.match(original_ms_key)
        if match:
            prefix_key_for_config = match.group(1).upper()  # Ключ для поиска в конфиге (AM, FR)
            version_number_part = match.group(2)

            template = config_data.get('MicroserviceVersions', {}).get(prefix_key_for_config)
            if template:
                # Извлекаем "чистое" имя сервиса из шаблона
                # Удаляем "(версия {{version}})" и лишние пробелы
                service_name_for_table = template.replace("{{version}}", "").replace("(версия )", "").strip()
            else:
                service_name_for_table = prefix_key_for_config  # Если шаблона нет, используем префикс
        else:
            logger.warning(
                f"Не удалось разобрать ключ микросервиса '{original_ms_key}' для сводной таблицы, будет использован как есть.")
            # В этом случае service_name_for_table = original_ms_key, version_number_part = ""

        # Ключ для проверки уникальности строки в таблице
        summary_tuple = (service_name_for_table, version_number_part)
        if summary_tuple not in seen_summary_entries:
            microservices_summary.append({
                'service_name': service_name_for_table,
                'version_number': version_number_part
            })
            seen_summary_entries.add(summary_tuple)
            logger.debug(
                f"Для сводной таблицы: Сервис='{service_name_for_table}', Версия='{version_number_part}' (из ключа '{original_ms_key}')")

    return microservices_summary


def main():
    setup_logging()
    logger = logging.getLogger(__name__)

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
    config_data = load_config(config_file_path)

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
    # ... (логирование параметров) ...

    raw_task_data, header_map, fix_versions_col_indices, issue_type_col_idx = \
        csv_importer.load_and_process_issues(csv_file_path, col_config)

    if raw_task_data is None:
        sys.exit(1)
    # ... (логирование количества прочитанных строк) ...

    global_version_part = csv_importer.find_global_version_title(raw_task_data, fix_versions_col_indices)
    # ... (логика формирования final_release_title) ...
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
        sys.exit(1)
    # ... (логирование количества сгруппированных версий) ...

    microservices_for_summary_table = []
    if grouped_issues_data:
        microservices_for_summary_table = extract_microservice_info_for_summary(
            grouped_issues_data.keys(),
            config_data
        )
        logger.info(f"Собрано {len(microservices_for_summary_table)} записей для сводной таблицы микросервисов.")
        logger.debug(f"Данные для сводной таблицы: {microservices_for_summary_table}")

    logger.info(f"Генерация DOCX документа '{docx_file_path}' для релиза '{final_release_title}'...")
    success = docx_creator.create_release_notes_docx(
        docx_file_path,
        final_release_title,
        grouped_issues_data,
        col_config['use_issue_type_grouping'],
        microservices_summary_data=microservices_for_summary_table,
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