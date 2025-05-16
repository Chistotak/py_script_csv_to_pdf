# release_notes_generator/csv_importer.py
import csv
import logging
import re
from collections import defaultdict

logger = logging.getLogger(__name__)


def sanitize_text_csv(text):
    if text is None: return ""
    return str(text).strip()


def extract_client_name(client_contract_string):
    if not client_contract_string:
        return "Не указан"
    return client_contract_string.split('#')[0].strip()


def load_and_process_issues(csv_filepath, col_config):
    all_rows = []
    header = []
    header_map = {}
    client_contract_col_index = None  # Инициализируем здесь

    try:
        with open(csv_filepath, mode='r', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile, delimiter=',')
            header = next(reader)
            logger.debug(f"CSV Headers: {header}")
            for i, col_name in enumerate(header):
                clean_col_name = sanitize_text_csv(col_name)
                if clean_col_name not in header_map:
                    header_map[clean_col_name] = i

            required_cols_present = True
            cols_to_check_existence = {  # Колонки, чье существование важно
                'key': col_config['key'],
                'customer_desc': col_config['customer_desc'],
                'install_instructions': col_config['install_instructions'],
                'fix_versions': col_config['fix_versions_name']  # Имя колонки, а не значение из col_config
            }
            if col_config.get('use_issue_type_grouping', False):
                cols_to_check_existence['issue_type'] = col_config['issue_type']
            if col_config.get('use_client_grouping', False):
                if 'client_contract' not in col_config or not col_config['client_contract']:
                    logger.error("Группировка по клиенту включена, но 'client_contract' не задан в [Columns] конфига.")
                    required_cols_present = False
                else:
                    cols_to_check_existence['client_contract'] = col_config['client_contract']

            for col_key_internal, col_name_in_csv in cols_to_check_existence.items():
                # Для fix_versions ищем имя в header, для остальных - в header_map
                if col_key_internal == 'fix_versions':
                    if col_name_in_csv not in header:
                        logger.error(
                            f"Ошибка: Колонка '{col_name_in_csv}' (для поля '{col_key_internal}') не найдена в CSV заголовках.")
                        required_cols_present = False
                elif col_name_in_csv not in header_map:
                    logger.error(
                        f"Ошибка: Колонка '{col_name_in_csv}' (для поля '{col_key_internal}') не найдена в CSV файле.")
                    required_cols_present = False

            issue_type_col_index = None
            if col_config.get('use_issue_type_grouping', False):
                issue_type_col_name = col_config.get('issue_type')
                if issue_type_col_name in header_map:
                    issue_type_col_index = header_map[issue_type_col_name]
                else:  # Уже должно быть поймано выше, но для безопасности
                    logger.warning(
                        f"Колонка типа задачи '{issue_type_col_name}' не найдена. Группировка по типу будет отключена.")
                    col_config['use_issue_type_grouping'] = False

            if col_config.get('use_client_grouping', False):
                client_contract_col_name = col_config.get('client_contract')
                if client_contract_col_name in header_map:
                    client_contract_col_index = header_map[client_contract_col_name]
                else:
                    logger.error(
                        f"Колонка клиента '{client_contract_col_name}' не найдена. Группировка по клиенту невозможна.")
                    col_config['use_client_grouping'] = False  # Отключаем, если колонка не найдена
                    # required_cols_present = False # Можно и так, если это критично

            if not required_cols_present:
                return None, None, None, None, None

            fix_versions_col_indices = [i for i, h_col in enumerate(header) if
                                        sanitize_text_csv(h_col) == col_config['fix_versions_name']]
            if not fix_versions_col_indices:
                logger.error(f"Крит. ошибка: Колонка '{col_config['fix_versions_name']}' не найдена.")
                return None, None, None, None, None

            for i, row in enumerate(reader):
                if len(row) == len(header):
                    all_rows.append([sanitize_text_csv(cell) for cell in row])
                elif any(cell.strip() for cell in row):
                    logger.warning(f"Строка {i + 2}: Пропуск...")

        logger.info(f"load_and_process_issues: Успешно прочитано {len(all_rows)} строк данных.")
        return all_rows, header_map, fix_versions_col_indices, issue_type_col_index, client_contract_col_index

    except FileNotFoundError:
        logger.error(f"Ошибка: CSV файл '{csv_filepath}' не найден.")
        return None, None, None, None, None
    except StopIteration:
        logger.error(f"Ошибка: CSV файл '{csv_filepath}' пуст или содержит только строку заголовков.")
        return None, None, None, None, None
    except Exception as e:
        logger.error(f"Ошибка при чтении CSV файла '{csv_filepath}': {e}", exc_info=True)
        return None, None, None, None, None


def group_issues(all_tasks_data, header_map, col_config,
                 fix_versions_col_indices, issue_type_col_index, client_contract_col_index,
                 main_config_data):
    use_client_grouping = col_config.get('use_client_grouping', False) and client_contract_col_index is not None
    use_type_grouping = col_config.get('use_issue_type_grouping', False) and issue_type_col_index is not None

    logger.info(f"Настройки группировки: по клиенту={use_client_grouping}, по типу задачи={use_type_grouping}")

    # Динамическое создание вложенности defaultdict
    if use_client_grouping and use_type_grouping:
        grouped_issues = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))  # ms -> client -> type -> tasks
    elif use_client_grouping:
        grouped_issues = defaultdict(lambda: defaultdict(list))  # ms -> client -> tasks
    elif use_type_grouping:
        grouped_issues = defaultdict(lambda: defaultdict(list))  # ms -> type -> tasks
    else:
        grouped_issues = defaultdict(list)  # ms -> tasks

    microservice_version_pattern = re.compile(r"^([A-Z]{2})(\d+(\.\d+){1,2})$")
    key_col_idx = header_map.get(col_config['key'])
    cust_desc_col_idx = header_map.get(col_config['customer_desc'])
    install_instr_col_idx = header_map.get(col_config['install_instructions'])

    if key_col_idx is None: return None

    for row_num, raw_row_data in enumerate(all_tasks_data):
        task_key_value = raw_row_data[key_col_idx] if key_col_idx < len(raw_row_data) else f"ROW_{row_num + 1}_NO_KEY"

        task_versions_from_row = []
        for index in fix_versions_col_indices:
            if index < len(raw_row_data) and raw_row_data[index]:
                task_versions_from_row.extend([v.strip() for v in raw_row_data[index].split(',') if v.strip()])
        task_versions_unique = list(set(task_versions_from_row))
        current_microservice_versions_original = [ver for ver in task_versions_unique if
                                                  microservice_version_pattern.match(ver)]
        if not current_microservice_versions_original: continue

        task_details = {
            'key': task_key_value,
            'cust_desc': raw_row_data[cust_desc_col_idx] if cust_desc_col_idx is not None and cust_desc_col_idx < len(
                raw_row_data) and raw_row_data[cust_desc_col_idx] else "",
            'install_instr': raw_row_data[
                install_instr_col_idx] if install_instr_col_idx is not None and install_instr_col_idx < len(
                raw_row_data) and raw_row_data[install_instr_col_idx] else ""
        }

        client_name_for_group = "Общие задачи"  # Используется если use_client_grouping = False
        if use_client_grouping:
            raw_client_string = raw_row_data[client_contract_col_index] if client_contract_col_index < len(
                raw_row_data) else ""
            client_name_for_group = extract_client_name(raw_client_string)
            logger.debug(
                f"Задача {task_key_value}: клиент '{client_name_for_group}' (из строки: '{raw_client_string[:50]}...')")

        issue_type_display_for_group = "Задачи"  # Используется если use_type_grouping = False
        if use_type_grouping:
            system_issue_type = raw_row_data[issue_type_col_index] if issue_type_col_index < len(
                raw_row_data) else "Не указан тип"
            system_issue_type = system_issue_type.strip() if system_issue_type else "Не указан тип"
            issue_type_display_for_group = main_config_data.get('IssueTypeNames', {}).get(system_issue_type,
                                                                                          system_issue_type)
            logger.debug(
                f"Задача {task_key_value}: тип '{issue_type_display_for_group}' (системный: '{system_issue_type}')")

        for ms_ver_key in current_microservice_versions_original:
            if use_client_grouping and use_type_grouping:
                grouped_issues[ms_ver_key][client_name_for_group][issue_type_display_for_group].append(task_details)
            elif use_client_grouping:
                grouped_issues[ms_ver_key][client_name_for_group].append(task_details)
            elif use_type_grouping:
                grouped_issues[ms_ver_key][issue_type_display_for_group].append(task_details)
            else:
                grouped_issues[ms_ver_key].append(task_details)

    # Сортировка
    for ms_ver_key in grouped_issues:
        if use_client_grouping and use_type_grouping:
            for client_key in grouped_issues[ms_ver_key]:
                for type_key in grouped_issues[ms_ver_key][client_key]:
                    grouped_issues[ms_ver_key][client_key][type_key].sort(key=lambda x: x['key'])
        elif use_client_grouping:
            for client_key in grouped_issues[ms_ver_key]:
                grouped_issues[ms_ver_key][client_key].sort(key=lambda x: x['key'])
        elif use_type_grouping:
            for type_key in grouped_issues[ms_ver_key]:
                grouped_issues[ms_ver_key][type_key].sort(key=lambda x: x['key'])
        else:
            grouped_issues[ms_ver_key].sort(key=lambda x: x['key'])

    logger.info(f"Группировка задач завершена.")
    return grouped_issues


def find_global_version_title(all_tasks_data, fix_versions_col_indices):
    # ... (код этой функции без изменений) ...
    global_versions_found = set()
    for task_data_row in all_tasks_data:
        current_task_versions = []
        for index in fix_versions_col_indices:
            if index < len(task_data_row) and task_data_row[index]:
                current_task_versions.extend([v.strip() for v in task_data_row[index].split(',')])
        for version in current_task_versions:
            if "(global)" in version:
                title = version.replace("(global)", "").strip()
                if title: global_versions_found.add(title)
    if not global_versions_found:
        logger.debug("Глобальная версия с суффиксом '(global)' не найдена.")
        return ""
    sorted_global_versions = sorted(list(global_versions_found))
    if len(sorted_global_versions) > 1:
        logger.warning(f"Найдено несколько глобальных версий: {sorted_global_versions}. Используется первая.")
    final_title_part = sorted_global_versions[0]
    logger.info(f"Найдена часть глобальной версии для заголовка: '{final_title_part}'")
    return final_title_part