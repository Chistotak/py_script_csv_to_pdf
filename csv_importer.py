# release_notes_generator/csv_importer.py

import csv
import logging
import re
from collections import defaultdict

logger = logging.getLogger(__name__)


def sanitize_text_csv(text):
    if text is None:
        return ""
    return str(text).strip()


def load_and_process_issues(csv_filepath, col_config):
    all_rows = []
    header = []
    header_map = {}

    try:
        with open(csv_filepath, mode='r', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile, delimiter=',')

            header = next(reader)
            logger.debug(f"CSV Headers: {header}")
            for i, col_name in enumerate(header):
                clean_col_name = sanitize_text_csv(col_name)
                if clean_col_name not in header_map:  # Берем первый индекс, если имена дублируются
                    header_map[clean_col_name] = i
            logger.debug(f"Header map (name to index): {header_map}")

            required_cols_present = True
            critical_cols_to_check = {
                'key': col_config['key'],
                'customer_desc': col_config['customer_desc'],
                'install_instructions': col_config['install_instructions']
                # fix_versions_name и issue_type проверяются отдельно
            }

            for col_key, col_val_name in critical_cols_to_check.items():
                if col_val_name not in header_map:
                    logger.error(
                        f"Ошибка: Обязательная колонка '{col_val_name}' (для поля '{col_key}') не найдена в CSV файле.")
                    required_cols_present = False

            if col_config['fix_versions_name'] not in header:  # Ищем точное совпадение в исходных заголовках
                logger.error(
                    f"Ошибка: Колонка '{col_config['fix_versions_name']}' (для версий) не найдена в CSV файле.")
                required_cols_present = False

            issue_type_col_index = None
            if col_config.get('use_issue_type_grouping', False):
                issue_type_col_name = col_config.get('issue_type')
                if issue_type_col_name and issue_type_col_name in header_map:
                    issue_type_col_index = header_map[issue_type_col_name]
                else:
                    logger.warning(
                        f"Колонка для типа задачи '{issue_type_col_name}' не найдена или не указана. Группировка по типу задачи будет отключена.")
                    col_config['use_issue_type_grouping'] = False

            if not required_cols_present:
                logger.error("Одна или несколько обязательных колонок отсутствуют. Прерывание загрузки.")
                return None, None, None, None

            fix_versions_col_indices = [i for i, h_col in enumerate(header) if
                                        sanitize_text_csv(h_col) == col_config['fix_versions_name']]
            if not fix_versions_col_indices:
                logger.error(
                    f"Критическая ошибка: Ни одной колонки с именем '{col_config['fix_versions_name']}' не найдено.")
                return None, None, None, None
            logger.debug(f"Индексы для колонки '{col_config['fix_versions_name']}': {fix_versions_col_indices}")

            for i, row in enumerate(reader):
                if len(row) == len(header):
                    all_rows.append([sanitize_text_csv(cell) for cell in row])
                elif any(cell.strip() for cell in row):  # Если строка не пустая, но не соответствует
                    logger.warning(
                        f"Строка {i + 2}: Пропуск строки с некорректным числом полей (ожидалось {len(header)}, получено {len(row)}): {str(row)[:100]}...")

        logger.info(f"load_and_process_issues: Успешно прочитано {len(all_rows)} строк данных.")
        logger.debug(f"load_and_process_issues: Header map: {header_map}")
        logger.debug(f"load_and_process_issues: Fix Version/s indices: {fix_versions_col_indices}")
        logger.debug(f"load_and_process_issues: Issue Type index: {issue_type_col_index}")
        return all_rows, header_map, fix_versions_col_indices, issue_type_col_index

    except FileNotFoundError:
        logger.error(f"Ошибка: CSV файл '{csv_filepath}' не найден.")
        return None, None, None, None
    except StopIteration:  # Файл пустой или содержит только заголовки
        logger.error(f"Ошибка: CSV файл '{csv_filepath}' пуст или содержит только строку заголовков.")
        return None, None, None, None
    except Exception as e:
        logger.error(f"Ошибка при чтении CSV файла '{csv_filepath}': {e}", exc_info=True)
        return None, None, None, None


def find_global_version_title(all_tasks_data, fix_versions_col_indices):
    global_versions_found = set()
    for task_data_row in all_tasks_data:
        current_task_versions = []
        for index in fix_versions_col_indices:
            if index < len(task_data_row) and task_data_row[index]:
                current_task_versions.extend([v.strip() for v in task_data_row[index].split(',')])

        for version in current_task_versions:
            if "(global)" in version:
                title = version.replace("(global)", "").strip()
                if title:  # Убедимся, что после удаления (global) что-то осталось
                    global_versions_found.add(title)

    if not global_versions_found:
        logger.warning(
            "Глобальная версия с суффиксом '(global)' не найдена. Используется стандартный заголовок 'Описание Релиза'.")
        return "Описание Релиза"

    sorted_global_versions = sorted(list(global_versions_found))
    if len(sorted_global_versions) > 1:
        logger.warning(
            f"Найдено несколько уникальных глобальных версий: {sorted_global_versions}. Используется первая: {sorted_global_versions[0]}")

    final_title = sorted_global_versions[0]
    logger.info(f"Найдена глобальная версия для заголовка: '{final_title}'")
    return final_title


def group_issues_by_version_and_type(all_tasks_data, header_map, col_config, fix_versions_col_indices,
                                     issue_type_col_index):
    use_type_grouping = col_config.get('use_issue_type_grouping', False) and issue_type_col_index is not None

    if use_type_grouping:
        grouped_issues = defaultdict(lambda: defaultdict(list))
    else:
        grouped_issues = defaultdict(list)

    microservice_version_pattern = re.compile(r"^([A-Z]{2}\d+(\.\d+){1,2})$")

    key_col_idx = header_map.get(col_config['key'])
    cust_desc_col_idx = header_map.get(col_config['customer_desc'])
    install_instr_col_idx = header_map.get(col_config['install_instructions'])

    if key_col_idx is None:
        logger.critical(
            f"Критическая ошибка: Индекс для ключа задачи '{col_config['key']}' не определен. Группировка невозможна.")
        return None

    tasks_processed_for_grouping = 0
    tasks_matched_microservice_version = 0

    for row_num, raw_row_data in enumerate(all_tasks_data):
        task_key_value = raw_row_data[key_col_idx] if key_col_idx < len(raw_row_data) else f"ROW_{row_num + 1}_NO_KEY"

        task_versions_from_row = []
        for index in fix_versions_col_indices:
            if index < len(raw_row_data) and raw_row_data[index]:
                task_versions_from_row.extend([v.strip() for v in raw_row_data[index].split(',') if v.strip()])

        task_versions_unique = list(set(task_versions_from_row))

        current_microservice_versions = [ver for ver in task_versions_unique if microservice_version_pattern.match(ver)]

        if not current_microservice_versions:
            logger.debug(
                f"Задача {task_key_value}: не найдено подходящих версий микросервисов из {task_versions_unique}. Пропуск.")
            continue

        tasks_matched_microservice_version += 1
        logger.debug(f"Задача {task_key_value}: найдены версии микросервисов: {current_microservice_versions}")

        task_details = {
            'key': task_key_value,
            'cust_desc': raw_row_data[cust_desc_col_idx] if cust_desc_col_idx is not None and cust_desc_col_idx < len(
                raw_row_data) and raw_row_data[cust_desc_col_idx] else "",
            'install_instr': raw_row_data[
                install_instr_col_idx] if install_instr_col_idx is not None and install_instr_col_idx < len(
                raw_row_data) and raw_row_data[install_instr_col_idx] else ""
        }

        current_issue_type = "Задачи"  # Тип по умолчанию
        if use_type_grouping:
            issue_type_val = raw_row_data[issue_type_col_index] if issue_type_col_index < len(raw_row_data) else None
            if issue_type_val:
                current_issue_type = issue_type_val.strip()
            else:
                current_issue_type = "Не указан тип"

        for ms_ver in current_microservice_versions:
            if use_type_grouping:
                grouped_issues[ms_ver][current_issue_type].append(task_details)
            else:
                grouped_issues[ms_ver].append(task_details)
            tasks_processed_for_grouping += 1

    logger.info(f"Группировка: обработано {tasks_matched_microservice_version} задач с версиями микросервисов.")
    logger.info(f"Группировка: всего записей о задачах добавлено в группы: {tasks_processed_for_grouping}.")

    # Сортировка
    for ms_ver in grouped_issues:
        if use_type_grouping:
            for i_type in grouped_issues[ms_ver]:
                grouped_issues[ms_ver][i_type].sort(key=lambda x: x['key'])
        else:
            grouped_issues[ms_ver].sort(key=lambda x: x['key'])

    return grouped_issues