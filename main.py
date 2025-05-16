# release_notes_generator/main.py
import argparse
import logging
import sys
import os
import configparser
import re

# Предполагается, что csv_importer.py и docx_creator.py находятся в той же директории
import csv_importer

try:
    import docx_creator
except ImportError:
    # Сообщение об ошибке выведет сам docx_creator при попытке импорта docx
    sys.exit(1)

# --- Стандартные значения, если ничего не найдено ни в CLI, ни в конфиге ---
DEFAULT_MAIN_CONFIG_FILE = "config.ini"
DEFAULT_STYLES_CONFIG_FILE = "styles.ini"
DEFAULT_COL_ISSUE_KEY = "Issue key"
DEFAULT_COL_FIX_VERSIONS_NAME = "Fix Version/s"
DEFAULT_COL_CUSTOMER_DESC = "Custom field (Description for the customer)"
DEFAULT_COL_INSTALL_INSTRUCTIONS = "Custom field (Инструкция по установке)"
DEFAULT_COL_ISSUE_TYPE = "Issue Type"
DEFAULT_COL_CLIENT_CONTRACT = "Custom field (Client\\Contract 1C)"  # Обратный слеш нужно экранировать или использовать raw string
DEFAULT_CSV_INPUT_FILE = "input.csv"  # Если даже в конфиге нет
DEFAULT_DOCX_OUTPUT_FILE = "output_releasenotes.docx"  # Если даже в конфиге нет


def setup_logging():
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(name)s - %(levelname)s - [%(funcName)s:%(lineno)d] - %(message)s',
                        handlers=[logging.StreamHandler(sys.stdout)])


def _parse_config_file(config_filepath, default_structure=None):
    """Вспомогательная функция для чтения одного ini файла."""
    logger = logging.getLogger(__name__)  # Получаем логгер для этой функции
    config = configparser.ConfigParser(allow_no_value=True, inline_comment_prefixes=('#', ';'))
    config.optionxform = str  # Сохраняем регистр ключей опций

    # Создаем копию default_structure, чтобы не изменять оригинал, если он был передан
    parsed_data = {section: dict(values) for section, values in default_structure.items()} if default_structure else {}

    if os.path.exists(config_filepath):
        try:
            config.read(config_filepath, encoding='utf-8')
            logger.info(f"Конфигурационный файл '{config_filepath}' успешно загружен.")
            for section_name_from_file in config.sections():
                file_values = dict(config.items(section_name_from_file))
                # Так как optionxform = str, имена секций в config.sections() будут такими же, как в файле
                # Мы обновляем или добавляем секцию в parsed_data
                if section_name_from_file in parsed_data:
                    parsed_data[section_name_from_file].update(file_values)
                else:
                    parsed_data[section_name_from_file] = file_values
        except configparser.Error as e:
            logger.error(f"Ошибка при чтении конфигурационного файла '{config_filepath}': {e}")
    else:
        logger.warning(
            f"Конфигурационный файл '{config_filepath}' не найден. Будут использованы значения по умолчанию (если определены в default_structure).")

    return parsed_data


def load_all_configs(main_config_path=DEFAULT_MAIN_CONFIG_FILE):
    """Загружает основной конфиг и, если указано, конфиг стилей."""
    logger = logging.getLogger(__name__)

    main_config_defaults = {
        'General': {
            'csv_input_file': DEFAULT_CSV_INPUT_FILE,
            'docx_output_file': DEFAULT_DOCX_OUTPUT_FILE,
            'logo_path': '',  # Путь к лого по умолчанию пустой
            'logo_width_inches': '1.5',
            'release_title_format': "{{global_version}}",
            'use_issue_type_grouping': 'true',
            'use_client_grouping': 'false',  # По умолчанию группировка по клиенту отключена
            'styles_config_file': DEFAULT_STYLES_CONFIG_FILE
        },
        'Columns': {
            'key': DEFAULT_COL_ISSUE_KEY,
            'fix_versions': DEFAULT_COL_FIX_VERSIONS_NAME,
            'customer_desc': DEFAULT_COL_CUSTOMER_DESC,
            'install_instructions': DEFAULT_COL_INSTALL_INSTRUCTIONS,
            'issue_type': DEFAULT_COL_ISSUE_TYPE,
            'client_contract': DEFAULT_COL_CLIENT_CONTRACT  # Добавляем дефолт для колонки клиента
        },
        'MicroserviceVersions': {},  # Пустые по умолчанию, должны быть в файле
        'IssueTypeNames': {}  # Пустые по умолчанию
    }
    main_config_data = _parse_config_file(main_config_path, main_config_defaults)
    # Сохраняем путь к директории основного конфига для разрешения относительных путей к другим файлам
    main_config_data['_config_dir_'] = os.path.dirname(os.path.abspath(main_config_path))
    logger.debug(f"Загружена основная конфигурация: {main_config_data}")

    # Дефолтная структура для конфига стилей (заполните по аналогии с вашим styles.ini)
    styles_config_defaults = {
        'Fonts': {'main': 'Arial', 'title': 'Calibri Light', 'section_header': 'Calibri', 'client_header': 'Calibri',
                  'issue_type_header': 'Calibri', 'task_key': 'Arial'},
        'FontSizes': {'title': '22', 'date': '10', 'summary_table_title': '14', 'summary_table_header': '11',
                      'summary_table_text': '10', 'ms_version_header': '16', 'client_header': '14',
                      'issue_type_header': '13', 'task_key': '11', 'task_description': '10',
                      'install_instruction_label': '10', 'install_instruction_text': '10', 'normal_style_base': '11'},
        'Colors': {'title': '003366', 'date_text': '595959', 'summary_table_title': '2F75B5',
                   'section_header': '2F75B5', 'client_header': '365F91', 'sub_header': '4A86E8', 'task_key': '000000',
                   'task_description': '333333', 'install_instruction_label': '1D1D1D',
                   'install_instruction_text': '4F4F4F', 'table_header_text': '000000', 'table_text': '000000'},
        'Spacing': {'after_title': '6', 'after_date': '24', 'after_summary_table_title': '8',
                    'after_summary_table': '18', 'after_ms_version_header': '6', 'after_client_header': '5',
                    'after_issue_type_header': '4', 'task_key_after': '1', 'task_description_after': '2',
                    'install_label_after': '1', 'install_text_after': '8', 'section_after_space': '12',
                    'normal_paragraph_after': '6'},
        'TableLayout': {'summary_table_col1_width_inches': '4.0', 'summary_table_col2_width_inches': '1.5'}
    }

    styles_file_path = main_config_data.get('General', {}).get('styles_config_file', DEFAULT_STYLES_CONFIG_FILE)
    # Если styles_file_path относительный, делаем его относительно директории основного конфига
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
    parser.add_argument("--styles-config", help="Конфиг стилей (переопред. значение из основного конфига)")
    parser.add_argument("--csv-file", help="Входной CSV (переопред. значение из основного конфига)")
    parser.add_argument("--docx-file", help="Выходной DOCX (переопред. значение из основного конфига)")

    # Аргументы для переопределения имен колонок
    parser.add_argument("--col-key", help=f"Переопределить имя колонки ключа задачи")
    parser.add_argument("--col-fix-versions", help=f"Переопределить имя колонки версий")
    parser.add_argument("--col-customer-desc", help=f"Переопределить имя колонки описания для клиента")
    parser.add_argument("--col-install-instructions", help=f"Переопределить имя колонки инструкции по установке")
    parser.add_argument("--col-issue-type", help=f"Переопределить имя колонки типа задачи")
    parser.add_argument("--col-client-contract", help="Переопределить имя колонки клиента/контракта")

    # Флаги для управления группировкой
    parser.add_argument("--no-issue-type-grouping", action='store_true', help="Отключить группировку по типу задачи")
    parser.add_argument("--no-client-grouping", action='store_true', help="Отключить группировку по клиенту")

    args = parser.parse_args()

    # --- Загрузка конфигураций ---
    main_cfg, styles_cfg = load_all_configs(args.config)  # Используем путь из CLI или дефолт

    # Если путь к конфигу стилей переопределен через CLI
    if args.styles_config:
        styles_cfg_path_cli = args.styles_config
        # Если путь относительный, он будет от текущей рабочей директории,
        # либо можно сделать его относительно директории основного конфига:
        # if not os.path.isabs(styles_cfg_path_cli):
        #    styles_cfg_path_cli = os.path.join(main_cfg['_config_dir_'], styles_cfg_path_cli)
        styles_cfg = _parse_config_file(styles_cfg_path_cli, styles_cfg)  # Перезагружаем с учетом предыдущих дефолтов
        logger.info(f"Конфигурация стилей перезагружена из CLI аргумента: {styles_cfg_path_cli}")

    # --- Определение параметров с учетом приоритетов: CLI > config.ini > дефолты в коде ---
    csv_fpath = args.csv_file if args.csv_file else main_cfg['General'].get('csv_input_file')
    docx_fpath = args.docx_file if args.docx_file else main_cfg['General'].get('docx_output_file')

    col_cfg = {
        'key': args.col_key if args.col_key else main_cfg['Columns'].get('key'),
        'fix_versions_name': args.col_fix_versions if args.col_fix_versions else main_cfg['Columns'].get(
            'fix_versions'),
        'customer_desc': args.col_customer_desc if args.col_customer_desc else main_cfg['Columns'].get('customer_desc'),
        'install_instructions': args.col_install_instructions if args.col_install_instructions else main_cfg[
            'Columns'].get('install_instructions'),
        'issue_type': args.col_issue_type if args.col_issue_type else main_cfg['Columns'].get('issue_type'),
        'client_contract': args.col_client_contract if args.col_client_contract else main_cfg['Columns'].get(
            'client_contract')
    }

    # Управление флагами группировки
    col_cfg['use_issue_type_grouping'] = not args.no_issue_type_grouping if args.no_issue_type_grouping \
        else main_cfg['General'].get('use_issue_type_grouping', 'true').lower() == 'true'
    col_cfg['use_client_grouping'] = not args.no_client_grouping if args.no_client_grouping \
        else main_cfg['General'].get('use_client_grouping', 'false').lower() == 'true'

    logger.info(f"--- Начало генерации отчета ---")
    logger.info(f"Основной конфиг: {os.path.abspath(args.config)}")
    # Определяем фактический путь к файлу стилей для логирования
    actual_styles_config_path = args.styles_config if args.styles_config else \
        os.path.join(main_cfg['_config_dir_'],
                     main_cfg['General'].get('styles_config_file', DEFAULT_STYLES_CONFIG_FILE))
    if args.styles_config and not os.path.isabs(args.styles_config):  # Если из CLI и относительный
        actual_styles_config_path = os.path.abspath(args.styles_config)  # То от текущей директории

    logger.info(f"Конфиг стилей: {os.path.abspath(actual_styles_config_path)}")
    logger.info(f"Входной CSV: {os.path.abspath(csv_fpath)}")
    logger.info(f"Выходной DOCX: {os.path.abspath(docx_fpath)}")
    logger.info(
        f"Настройки группировки: по клиенту={col_cfg['use_client_grouping']}, по типу={col_cfg['use_issue_type_grouping']}")

    raw_task_data, header_map, fix_versions_col_indices, issue_type_col_idx, client_contract_col_idx = \
        csv_importer.load_and_process_issues(csv_fpath, col_cfg)
    if raw_task_data is None: sys.exit(1)

    global_version_part = csv_importer.find_global_version_title(raw_task_data, fix_versions_col_indices)
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

    logger.info("Группировка задач...")
    grouped_issues_data = csv_importer.group_issues(
        raw_task_data, header_map, col_cfg,
        fix_versions_col_indices, issue_type_col_idx, client_contract_col_idx,
        main_cfg  # Передаем основной конфиг для маппинга IssueTypeNames
    )
    if grouped_issues_data is None: sys.exit(1)

    ms_summary_data = []
    if grouped_issues_data:
        ms_summary_data = docx_creator.extract_microservice_info_for_summary_table(grouped_issues_data.keys(), main_cfg)

    logger.info(f"Генерация DOCX: '{docx_fpath}' для релиза '{final_release_title}'...")
    success = docx_creator.create_release_notes_docx(
        docx_fpath, final_release_title, grouped_issues_data,
        col_cfg['use_client_grouping'],
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