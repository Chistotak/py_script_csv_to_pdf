# release_notes_generator/main.py

import argparse
import logging
import sys
import os
import configparser
import re

import csv_importer  # Предполагается, что он актуален

try:
    import docx_creator  # Обновленный docx_creator
except ImportError:
    sys.exit(1)

# --- Стандартные значения ---
DEFAULT_MAIN_CONFIG_FILE = "config.ini"
DEFAULT_STYLES_CONFIG_FILE = "styles.ini"
DEFAULT_COL_ISSUE_KEY = "Issue key";
DEFAULT_COL_FIX_VERSIONS_NAME = "Fix Version/s"
DEFAULT_COL_CUSTOMER_DESC = "Custom field (Description for the customer)";
DEFAULT_COL_INSTALL_INSTRUCTIONS = "Custom field (Инструкция по установке)"
DEFAULT_COL_ISSUE_TYPE = "Issue Type";
DEFAULT_CSV_INPUT_FILE = "input.csv"
DEFAULT_DOCX_OUTPUT_FILE = "output_releasenotes.docx"


def setup_logging():
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(name)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s',
                        handlers=[logging.StreamHandler(sys.stdout)])


def _parse_config_file(config_filepath, default_structure=None):
    logger = logging.getLogger(__name__)
    config = configparser.ConfigParser(allow_no_value=True, inline_comment_prefixes=('#', ';'))
    config.optionxform = str
    parsed_data = {section: dict(values) for section, values in default_structure.items()} if default_structure else {}
    if os.path.exists(config_filepath):
        try:
            config.read(config_filepath, encoding='utf-8')
            logger.info(f"Конфигурационный файл '{config_filepath}' успешно загружен.")
            for section_name_from_file in config.sections():
                file_values = dict(config.items(section_name_from_file))
                matched_section_key = section_name_from_file  # Так как optionxform = str, имена секций тоже сохраняют регистр
                if matched_section_key in parsed_data:
                    parsed_data[matched_section_key].update(file_values)
                else:
                    parsed_data[matched_section_key] = file_values
        except configparser.Error as e:
            logger.error(f"Ошибка при чтении '{config_filepath}': {e}")
    else:
        logger.warning(f"Файл '{config_filepath}' не найден. Используются дефолты (если есть).")
    return parsed_data


def load_all_configs(main_config_path=DEFAULT_MAIN_CONFIG_FILE):
    logger = logging.getLogger(__name__)
    main_config_defaults = {
        'General': {'csv_input_file': DEFAULT_CSV_INPUT_FILE, 'docx_output_file': DEFAULT_DOCX_OUTPUT_FILE,
                    'logo_path': '', 'logo_width_inches': '1.5', 'release_title_format': "{{global_version}}",
                    'use_issue_type_grouping': 'true', 'styles_config_file': DEFAULT_STYLES_CONFIG_FILE},
        'Columns': {'key': DEFAULT_COL_ISSUE_KEY, 'fix_versions': DEFAULT_COL_FIX_VERSIONS_NAME,
                    'customer_desc': DEFAULT_COL_CUSTOMER_DESC,
                    'install_instructions': DEFAULT_COL_INSTALL_INSTRUCTIONS,
                    'issue_type': DEFAULT_COL_ISSUE_TYPE},
        'MicroserviceVersions': {}, 'IssueTypeNames': {}
    }
    main_config_data = _parse_config_file(main_config_path, main_config_defaults)
    main_config_data['_config_dir_'] = os.path.dirname(
        os.path.abspath(main_config_path))  # Сохраняем путь к директории основного конфига
    logger.debug(f"Загружена основная конфигурация: {main_config_data}")

    styles_config_defaults = {
        'Fonts': {'main': 'Arial', 'title': 'Calibri Light'}, 'FontSizes': {'title': '22', 'normal_style_base': '11'},
        'Colors': {'title': '003366', 'normal_text': '333333'},
        'Spacing': {'after_title': '6', 'normal_paragraph_after': '6'},
        'TableLayout': {'summary_table_col1_width_inches': '4.0', 'summary_table_col2_width_inches': '1.5'}
    }
    styles_file_path = main_config_data.get('General', {}).get('styles_config_file', DEFAULT_STYLES_CONFIG_FILE)
    if not os.path.isabs(styles_file_path):
        styles_file_path = os.path.join(main_config_data['_config_dir_'], styles_file_path)

    styles_config_data = _parse_config_file(styles_file_path, styles_config_defaults)
    logger.debug(f"Загружена конфигурация стилей: {styles_config_data}")
    return main_config_data, styles_config_data


def main():
    setup_logging()
    logger = logging.getLogger(__name__)

    parser = argparse.ArgumentParser(description="Генератор DOCX релиза из JIRA CSV.")
    parser.add_argument("--config", default=DEFAULT_MAIN_CONFIG_FILE,
                        help=f"Основной конфиг (по умолч: {DEFAULT_MAIN_CONFIG_FILE})")
    parser.add_argument("--styles-config", help="Конфиг стилей (переопред. из основного)")
    parser.add_argument("--csv-file", help="Входной CSV (переопред. из основного конфига)")
    parser.add_argument("--docx-file", help="Выходной DOCX (переопред. из основного конфига)")
    parser.add_argument("--col-key", help="Колонка ключа задачи")
    parser.add_argument("--col-fix-versions", help="Колонка версий")
    parser.add_argument("--no-issue-type-grouping", action='store_true', help="Отключить группировку по типу")
    args = parser.parse_args()

    main_cfg, styles_cfg = load_all_configs(args.config)
    if args.styles_config:
        styles_cfg_path = args.styles_config
        if not os.path.isabs(
                styles_cfg_path):  # Делаем путь к конфигу стилей абсолютным от директории основного конфига
            styles_cfg_path = os.path.join(main_cfg['_config_dir_'], styles_cfg_path)
        styles_cfg = _parse_config_file(styles_cfg_path, styles_cfg)  # Перезагружаем с учетом предыдущих дефолтов
        logger.info(f"Конфигурация стилей перезагружена из CLI: {styles_cfg_path}")

    csv_fpath = args.csv_file if args.csv_file else main_cfg['General'].get('csv_input_file')
    docx_fpath = args.docx_file if args.docx_file else main_cfg['General'].get('docx_output_file')

    col_cfg = {
        'key': args.col_key if args.col_key else main_cfg['Columns'].get('key'),
        'fix_versions_name': args.col_fix_versions if args.col_fix_versions else main_cfg['Columns'].get(
            'fix_versions'),
        'customer_desc': main_cfg['Columns'].get('customer_desc'),
        'install_instructions': main_cfg['Columns'].get('install_instructions'),
        'issue_type': main_cfg['Columns'].get('issue_type'),
    }
    col_cfg['use_issue_type_grouping'] = not args.no_issue_type_grouping if args.no_issue_type_grouping else \
        main_cfg['General'].get('use_issue_type_grouping', 'true').lower() == 'true'

    logger.info(f"--- Начало генерации отчета ---")
    logger.info(f"Основной конфиг: {os.path.abspath(args.config)}")
    logger.info(
        f"Конфиг стилей: {os.path.abspath(styles_cfg_path if args.styles_config else main_cfg.get('General', {}).get('styles_config_file', 'styles.ini'))}")  # Показываем используемый путь
    logger.info(f"Входной CSV: {os.path.abspath(csv_fpath)}")
    logger.info(f"Выходной DOCX: {os.path.abspath(docx_fpath)}")

    raw_task_data, header_map, fix_versions_col_indices, issue_type_col_idx = \
        csv_importer.load_and_process_issues(csv_fpath, col_cfg)
    if raw_task_data is None: sys.exit(1)

    global_version_part = csv_importer.find_global_version_title(raw_task_data, fix_versions_col_indices)
    # ... (логика final_release_title как в предыдущей версии main.py, используя main_cfg['General']) ...
    release_title_override = main_cfg['General'].get('release_title_override')
    if release_title_override:
        final_release_title = release_title_override
    else:
        title_format_template = main_cfg['General'].get('release_title_format', "{{global_version}}")
        if global_version_part:
            final_release_title = title_format_template.replace("{{global_version}}", global_version_part)
        else:
            default_title_if_no_global = "Описание Релиза"
            if title_format_template and title_format_template != "{{global_version}}":
                final_release_title = title_format_template.replace("{{global_version}}", "Не указана").strip().rstrip(
                    ':').strip()
                if not final_release_title or final_release_title == main_cfg['General'].get('release_title_format',
                                                                                             "").replace(
                        "{{global_version}}", "").strip().rstrip(':').strip():
                    final_release_title = default_title_if_no_global
            else:
                final_release_title = default_title_if_no_global
            logger.warning(f"Глобальная версия не найдена, используется '{final_release_title}'")
    logger.info(f"Финальный заголовок: '{final_release_title}'")

    grouped_issues_data = csv_importer.group_issues_by_version_and_type(raw_task_data, header_map, col_cfg,
                                                                        fix_versions_col_indices, issue_type_col_idx,
                                                                        main_cfg)
    if grouped_issues_data is None: sys.exit(1)

    ms_summary_data = []
    if grouped_issues_data:
        ms_summary_data = docx_creator.extract_microservice_info_for_summary_table(grouped_issues_data.keys(), main_cfg)

    logger.info(f"Генерация DOCX: '{docx_fpath}' для релиза '{final_release_title}'...")
    success = docx_creator.create_release_notes_docx(
        docx_fpath, final_release_title, grouped_issues_data,
        col_cfg['use_issue_type_grouping'],
        microservices_summary_data=ms_summary_data,
        main_config=main_cfg, style_config=styles_cfg
    )

    if success:
        logger.info(f"--- Генерация отчета успешно завершена: {os.path.abspath(docx_fpath)} ---")
    else:
        logger.error("--- Ошибки при создании DOCX. ---"); sys.exit(1)


if __name__ == "__main__":
    main()