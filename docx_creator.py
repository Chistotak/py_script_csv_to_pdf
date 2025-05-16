# release_notes_generator/docx_creator.py

import logging
from datetime import datetime
import os
import re

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    # from docx.enum.table import WD_TABLE_ALIGNMENT # Если понадобится для выравнивания всей таблицы
except ImportError:
    logging.critical("Библиотека python-docx не установлена. Установите ее командой: pip install python-docx")
    raise

logger = logging.getLogger(__name__)


def sanitize_text_docx(text):
    if text is None:
        return ""
    return str(text)


def _add_formatted_paragraph(document, text, font_size=None, bold=False, italic=False, alignment=None,
                             left_indent_inches=None, font_name='Arial'):
    p = document.add_paragraph()
    if alignment:
        p.alignment = alignment
    if left_indent_inches:
        p.paragraph_format.left_indent = Inches(left_indent_inches)

    # Если текст пустой, добавляем пустой run, чтобы параграф не исчез при некоторых операциях
    run_text = sanitize_text_docx(text) if text else " "
    if text is None or text == "":  # Если исходный текст пустой, делаем run невидимым (пробел и маленький шрифт)
        if not text:  # если текст изначально был пустой
            run_text = " "  # нужен хоть какой-то символ для run
            font_size_actual = Pt(1)  # делаем пробел почти невидимым
    else:
        font_size_actual = Pt(font_size) if font_size else None

    run = p.add_run(run_text)

    if font_name:
        try:
            run.font.name = font_name
        except Exception as e:
            logger.warning(f"Не удалось установить шрифт '{font_name}' для run: {e}")

    if font_size_actual:
        run.font.size = font_size_actual

    if text is None or text == "":  # если текст был пустой, убираем жирность/курсив
        run.bold = False
        run.italic = False
    else:
        run.bold = bold
        run.italic = italic
    return p


def create_release_notes_docx(output_filename, title, grouped_data, use_issue_type_grouping,
                              microservices_summary_data=None,
                              default_font_name='Arial', config_data=None):
    document = Document()
    logger.info(f"Создание DOCX документа: {output_filename}")

    microservice_version_pattern_docx = re.compile(
        r"^([A-Z]{2})(\d+(\.\d+){1,2})$")  # Паттерн для разбора ключей grouped_data

    try:
        normal_style = document.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = default_font_name
        normal_font.size = Pt(10)
        logger.info(f"Попытка установить шрифт '{default_font_name}' Pt(10) для стиля 'Normal'.")
    except Exception as e:
        logger.warning(f"Не удалось установить шрифт по умолчанию для стиля 'Normal': {e}.")

    if config_data:
        logo_path_key = config_data.get('General', {}).get('logo_path')  # Получаем значение ключа
        if logo_path_key:  # Если ключ есть и не пустой
            # Преобразуем относительный путь в абсолютный, если он не абсолютный
            # Это полезно, если скрипт запускается не из той директории, где лежит config.ini или main.py
            # Предполагаем, что относительный путь в config.ini указан относительно директории, где лежит config.ini или main.py
            # Для простоты, пока считаем, что если путь относительный, то он от текущей рабочей директории скрипта
            # или используйте os.path.join(os.path.dirname(config_file_path_from_main), logo_path_key) если передавать путь к конфигу

            # Простой вариант: если путь не абсолютный, считаем его от текущей рабочей директории
            actual_logo_path = logo_path_key
            if not os.path.isabs(logo_path_key):
                # Можно сделать его относительно директории скрипта main.py
                # script_dir = os.path.dirname(os.path.abspath(sys.argv[0] или __file__ из main))
                # actual_logo_path = os.path.join(script_dir, logo_path_key)
                # Пока оставим как есть, предполагая, что путь корректен относительно места запуска
                pass

            if os.path.exists(actual_logo_path):
                try:
                    logo_width_str = config_data.get('General', {}).get('logo_width_inches', '1.5')
                    logo_width = float(logo_width_str)
                    p_logo = document.add_paragraph()
                    p_logo.add_run().add_picture(actual_logo_path, width=Inches(logo_width))
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    logger.info(f"Логотип '{actual_logo_path}' добавлен с шириной {logo_width} дюймов.")
                except ValueError:
                    logger.error(
                        f"Некорректное значение для 'logo_width_inches': '{logo_width_str}'. Логотип не добавлен.")
                except Exception as e:
                    logger.error(f"Не удалось добавить логотип '{actual_logo_path}': {e}")
            else:
                logger.warning(
                    f"Файл логотипа '{actual_logo_path}' (из ключа '{logo_path_key}') не найден. Логотип не добавлен.")
        else:
            logger.debug("Путь к логотипу (logo_path) не указан или пуст в конфигурации.")

    _add_formatted_paragraph(document, title, font_size=18, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                             font_name=default_font_name)
    _add_formatted_paragraph(document, f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M')}", font_size=10,
                             alignment=WD_ALIGN_PARAGRAPH.CENTER, font_name=default_font_name)

    if microservices_summary_data:
        document.add_paragraph()
        _add_formatted_paragraph(document, "Состав релиза по микросервисам:", font_size=14, bold=True,
                                 font_name=default_font_name)

        if microservices_summary_data:  # Проверяем, что список не пустой
            table = document.add_table(rows=1, cols=2)
            table.style = 'Table Grid'

            hdr_cells = table.rows[0].cells
            col1_hdr_text = 'Микросервис'
            col2_hdr_text = 'Версия'

            # Добавляем текст и форматируем заголовки
            p = hdr_cells[0].paragraphs[0]
            run = p.add_run(col1_hdr_text)
            run.bold = True
            run.font.name = default_font_name
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # или CENTER

            p = hdr_cells[1].paragraphs[0]
            run = p.add_run(col2_hdr_text)
            run.bold = True
            run.font.name = default_font_name
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # или CENTER

            for item in microservices_summary_data:
                row_cells = table.add_row().cells
                row_cells[0].text = item.get('service_name', 'N/A')
                row_cells[1].text = item.get('version_number', 'N/A')
                for i, cell in enumerate(row_cells):
                    if cell.paragraphs:
                        # Устанавливаем шрифт для каждой ячейки данных
                        p_cell = cell.paragraphs[0]
                        # Нужно добавить run, если его нет, или получить существующий
                        run_cell = p_cell.runs[0] if p_cell.runs else p_cell.add_run(cell.text)
                        # Если текст был добавлен через cell.text, то run уже есть, но может быть пустым
                        # Это более надежный способ получить или создать run с текстом:
                        if not p_cell.runs: p_cell.add_run(cell.text)
                        p_cell.runs[0].font.name = default_font_name
                        # p_cell.alignment = WD_ALIGN_PARAGRAPH.LEFT # если нужно выравнивание для данных

            logger.info(f"Таблица микросервисов ({len(microservices_summary_data)} строк) добавлена.")
        else:
            _add_formatted_paragraph(document, "Информация о версиях микросервисов в данном релизе отсутствует.",
                                     font_size=10, italic=True, font_name=default_font_name)
        document.add_paragraph()
    elif grouped_data:
        logger.info("Данные для сводной таблицы микросервисов отсутствуют или не были переданы.")
        document.add_paragraph()

    if not grouped_data:
        logger.info("Нет сгруппированных данных по задачам для добавления в DOCX.")
        if not microservices_summary_data:  # Если и таблицы не было
            _add_formatted_paragraph(document, "Нет задач для отображения в этом релизе.", font_name=default_font_name)
    else:
        sorted_ms_versions_original = sorted(grouped_data.keys())
        logger.info(
            f"Будет обработано {len(sorted_ms_versions_original)} версий микросервисов (детализация задач) для DOCX.")

        for ms_version_original_key in sorted_ms_versions_original:
            display_ms_version = ms_version_original_key
            if config_data:
                match = microservice_version_pattern_docx.match(ms_version_original_key)
                if match:
                    prefix_to_lookup = match.group(1).upper()
                    version_number = match.group(2)
                    template = config_data.get('MicroserviceVersions', {}).get(prefix_to_lookup)
                    if template:
                        display_ms_version = template.replace("{{version}}", version_number)
                    logger.debug(
                        f"Обработка версии: оригинал='{ms_version_original_key}', префикс для поиска='{prefix_to_lookup}', шаблон из конфига='{template}', отображаемое имя='{display_ms_version}'")
                else:
                    logger.warning(
                        f"Не удалось разобрать оригинальную версию микросервиса '{ms_version_original_key}' паттерном для отображения полного имени.")

            _add_formatted_paragraph(document, display_ms_version, font_size=16, bold=True, font_name=default_font_name)
            data_for_ms = grouped_data[ms_version_original_key]

            current_base_indent_main = 0.0  # Для отладки, в коде ниже он переопределяется
            if use_issue_type_grouping and isinstance(data_for_ms, dict):
                sorted_issue_types_display = sorted(data_for_ms.keys())
                for issue_type_display_name in sorted_issue_types_display:
                    _add_formatted_paragraph(document, issue_type_display_name, font_size=14, bold=True,
                                             left_indent_inches=0.25, font_name=default_font_name)
                    current_base_indent_main = 0.50
                    tasks_to_process = data_for_ms[issue_type_display_name]

                    for task_num, task in enumerate(tasks_to_process):
                        logger.debug(
                            f"Добавление задачи {task.get('key', 'N/A')} (тип: {issue_type_display_name}, версия: {display_ms_version})")
                        _add_formatted_paragraph(document, task['key'] + ":", bold=True,
                                                 left_indent_inches=current_base_indent_main,
                                                 font_name=default_font_name, font_size=10)

                        cust_desc_text = sanitize_text_docx(task['cust_desc'])
                        is_cust_desc_empty = not bool(task['cust_desc'])
                        _add_formatted_paragraph(document,
                                                 cust_desc_text if not is_cust_desc_empty else "Описание для клиента отсутствует.",
                                                 italic=is_cust_desc_empty,
                                                 left_indent_inches=current_base_indent_main,
                                                 font_name=default_font_name, font_size=10)

                        if task['install_instr']:
                            _add_formatted_paragraph(document, "Инструкция по установке:", bold=True,
                                                     left_indent_inches=current_base_indent_main,
                                                     font_name=default_font_name, font_size=10)
                            _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                                     left_indent_inches=current_base_indent_main,
                                                     font_name=default_font_name, font_size=10)

                        if task_num < len(tasks_to_process) - 1:
                            _add_formatted_paragraph(document,
                                                     text=None)  # Используем пустой текст для создания отступа
            elif isinstance(data_for_ms, list):
                current_base_indent_main = 0.25
                tasks_to_process = data_for_ms
                for task_num, task in enumerate(tasks_to_process):
                    logger.debug(f"Добавление задачи {task.get('key', 'N/A')} (версия: {display_ms_version})")
                    _add_formatted_paragraph(document, task['key'] + ":", bold=True,
                                             left_indent_inches=current_base_indent_main, font_name=default_font_name,
                                             font_size=10)
                    cust_desc_text = sanitize_text_docx(task['cust_desc'])
                    is_cust_desc_empty = not bool(task['cust_desc'])
                    _add_formatted_paragraph(document,
                                             cust_desc_text if not is_cust_desc_empty else "Описание для клиента отсутствует.",
                                             italic=is_cust_desc_empty, left_indent_inches=current_base_indent_main,
                                             font_name=default_font_name, font_size=10)
                    if task['install_instr']:
                        _add_formatted_paragraph(document, "Инструкция по установке:", bold=True,
                                                 left_indent_inches=current_base_indent_main,
                                                 font_name=default_font_name, font_size=10)
                        _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                                 left_indent_inches=current_base_indent_main,
                                                 font_name=default_font_name, font_size=10)
                    if task_num < len(tasks_to_process) - 1:
                        _add_formatted_paragraph(document, text=None)
            _add_formatted_paragraph(document, text=None)  # Отступ после секции микросервиса

    try:
        document.save(output_filename)
        logger.info(f"DOCX '{output_filename}' успешно сохранен.")
        return True
    except Exception as e:
        logger.error(f"Ошибка при сохранении DOCX '{output_filename}': {e}", exc_info=True)
        return False