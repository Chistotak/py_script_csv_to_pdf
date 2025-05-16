# release_notes_generator/docx_creator.py

import logging
from datetime import datetime
import os
import re  # Для разбора версий микросервисов

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
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

    run = p.add_run(sanitize_text_docx(text))
    if font_name:
        try:
            run.font.name = font_name
        except Exception as e:
            logger.warning(f"Не удалось установить шрифт '{font_name}' для run: {e}")
    if font_size:
        run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    return p


def create_release_notes_docx(output_filename, title, grouped_data, use_issue_type_grouping, default_font_name='Arial',
                              config_data=None):
    document = Document()
    logger.info(f"Создание DOCX документа: {output_filename}")

    # Паттерн для извлечения префикса и номера из ОРИГИНАЛЬНОГО ключа версии микросервиса
    microservice_version_pattern_docx = re.compile(r"^([A-Z]{2})(\d+(\.\d+){1,2})$")

    try:
        normal_style = document.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = default_font_name
        normal_font.size = Pt(10)
        logger.info(f"Попытка установить шрифт '{default_font_name}' Pt(10) для стиля 'Normal'.")
    except Exception as e:
        logger.warning(
            f"Не удалось установить шрифт по умолчанию для стиля 'Normal': {e}. Word будет использовать свои настройки.")

    # Добавление логотипа
    if config_data:
        logo_path = config_data.get('General', {}).get('logo_path')
        if logo_path:  # Проверяем, есть ли вообще ключ
            if os.path.exists(logo_path):  # Проверяем существование файла
                try:
                    logo_width_str = config_data.get('General', {}).get('logo_width_inches', '1.5')
                    logo_width = float(logo_width_str)
                    # Добавляем логотип, затем выравниваем параграф, в котором он оказался
                    p_logo = document.add_paragraph()
                    p_logo.add_run().add_picture(logo_path, width=Inches(logo_width))
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    logger.info(f"Логотип '{logo_path}' добавлен с шириной {logo_width} дюймов.")
                except ValueError:
                    logger.error(
                        f"Некорректное значение для 'logo_width_inches': '{logo_width_str}'. Логотип не добавлен.")
                except Exception as e:
                    logger.error(f"Не удалось добавить логотип '{logo_path}': {e}")
            else:
                logger.warning(f"Файл логотипа '{logo_path}' не найден. Логотип не добавлен.")
        else:
            logger.debug("Путь к логотипу (logo_path) не указан в конфигурации.")

    _add_formatted_paragraph(document, title, font_size=18, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                             font_name=default_font_name)
    _add_formatted_paragraph(document, f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M')}", font_size=10,
                             alignment=WD_ALIGN_PARAGRAPH.CENTER, font_name=default_font_name)
    document.add_paragraph()

    if not grouped_data:
        logger.info("Нет сгруппированных данных для добавления в DOCX.")
        _add_formatted_paragraph(document, "Нет задач для отображения в этом релизе.", font_name=default_font_name)
    else:
        sorted_ms_versions_original = sorted(grouped_data.keys())  # Ключи - это оригинальные "FR2.3.3"
        logger.info(f"Будет обработано {len(sorted_ms_versions_original)} версий микросервисов для DOCX.")

        for ms_version_original_key in sorted_ms_versions_original:
            # Получаем отображаемое имя версии микросервиса
            display_ms_version = ms_version_original_key  # По умолчанию
            if config_data:
                match = microservice_version_pattern_docx.match(ms_version_original_key)
                if match:
                    prefix = match.group(1).upper()
                    version_number = match.group(2)
                    template = config_data.get('MicroserviceVersions', {}).get(prefix)
                    if template:
                        display_ms_version = template.replace("{{version}}", version_number)

            _add_formatted_paragraph(document, display_ms_version, font_size=16, bold=True, font_name=default_font_name)
            data_for_ms = grouped_data[ms_version_original_key]  # Получаем данные по оригинальному ключу

            current_base_indent = 0.0
            if use_issue_type_grouping and isinstance(data_for_ms, dict):
                # Ключи data_for_ms - это уже отображаемые имена типов задач
                sorted_issue_types_display = sorted(data_for_ms.keys())
                for issue_type_display_name in sorted_issue_types_display:
                    _add_formatted_paragraph(document, issue_type_display_name, font_size=14, bold=True,
                                             left_indent_inches=0.25, font_name=default_font_name)
                    current_base_indent = 0.50
                    tasks_to_process = data_for_ms[issue_type_display_name]

                    for task_num, task in enumerate(tasks_to_process):
                        logger.debug(
                            f"Добавление задачи {task.get('key', 'N/A')} (тип: {issue_type_display_name}, версия: {display_ms_version})")
                        _add_formatted_paragraph(document, task['key'] + ":", bold=True,
                                                 left_indent_inches=current_base_indent, font_name=default_font_name,
                                                 font_size=10)

                        cust_desc_text = sanitize_text_docx(task['cust_desc'])
                        is_cust_desc_empty = not bool(task['cust_desc'])
                        _add_formatted_paragraph(document,
                                                 cust_desc_text if not is_cust_desc_empty else "Описание для клиента отсутствует.",
                                                 italic=is_cust_desc_empty,
                                                 left_indent_inches=current_base_indent, font_name=default_font_name,
                                                 font_size=10)

                        if task['install_instr']:
                            _add_formatted_paragraph(document, "Инструкция по установке:", bold=True,
                                                     left_indent_inches=current_base_indent,
                                                     font_name=default_font_name, font_size=10)
                            _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                                     left_indent_inches=current_base_indent,
                                                     font_name=default_font_name, font_size=10)

                        if task_num < len(tasks_to_process) - 1:
                            document.add_paragraph()
            elif isinstance(data_for_ms, list):  # Без группировки по типу
                current_base_indent = 0.25
                tasks_to_process = data_for_ms
                for task_num, task in enumerate(tasks_to_process):
                    logger.debug(f"Добавление задачи {task.get('key', 'N/A')} (версия: {display_ms_version})")
                    _add_formatted_paragraph(document, task['key'] + ":", bold=True,
                                             left_indent_inches=current_base_indent, font_name=default_font_name,
                                             font_size=10)
                    cust_desc_text = sanitize_text_docx(task['cust_desc'])
                    is_cust_desc_empty = not bool(task['cust_desc'])
                    _add_formatted_paragraph(document,
                                             cust_desc_text if not is_cust_desc_empty else "Описание для клиента отсутствует.",
                                             italic=is_cust_desc_empty, left_indent_inches=current_base_indent,
                                             font_name=default_font_name, font_size=10)
                    if task['install_instr']:
                        _add_formatted_paragraph(document, "Инструкция по установке:", bold=True,
                                                 left_indent_inches=current_base_indent, font_name=default_font_name,
                                                 font_size=10)
                        _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                                 left_indent_inches=current_base_indent, font_name=default_font_name,
                                                 font_size=10)
                    if task_num < len(tasks_to_process) - 1:
                        document.add_paragraph()
            document.add_paragraph()

    try:
        document.save(output_filename)
        logger.info(f"DOCX '{output_filename}' успешно сохранен.")
        return True
    except Exception as e:
        logger.error(f"Ошибка при сохранении DOCX '{output_filename}': {e}", exc_info=True)
        return False