# release_notes_generator/docx_creator.py

import logging
from datetime import datetime

try:
    from docx import Document
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    # from docx.enum.style import WD_STYLE_TYPE # Пока не используем явно
except ImportError:
    logging.critical("Библиотека python-docx не установлена. Установите ее командой: pip install python-docx")
    raise

logger = logging.getLogger(__name__)


def sanitize_text_docx(text):
    if text is None:
        return ""
    return str(text)  # Для DOCX базовая обработка может быть проще


def _add_formatted_paragraph(document, text, font_size=None, bold=False, italic=False, alignment=None,
                             left_indent_inches=None, font_name='Arial'):
    p = document.add_paragraph()
    if alignment:
        p.alignment = alignment
    if left_indent_inches:
        p.paragraph_format.left_indent = Inches(left_indent_inches)

    run = p.add_run(sanitize_text_docx(text))
    if font_name:  # Попытка установить шрифт
        try:
            run.font.name = font_name
        except Exception as e:
            logger.warning(f"Не удалось установить шрифт '{font_name}' для run: {e}")
    if font_size:
        run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    return p


def create_release_notes_docx(output_filename, title, grouped_data, use_issue_type_grouping, default_font_name='Arial'):
    document = Document()
    logger.info(f"Создание DOCX документа: {output_filename}")

    try:
        # Попытка установить шрифт по умолчанию для стиля 'Normal'
        # Это может не сработать, если шаблон по умолчанию не позволяет менять его так легко,
        # или если такого стиля нет (маловероятно для 'Normal').
        normal_style = document.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = default_font_name
        normal_font.size = Pt(10)  # Базовый размер
        logger.info(f"Попытка установить шрифт '{default_font_name}' для стиля 'Normal'.")
    except Exception as e:
        logger.warning(
            f"Не удалось установить шрифт по умолчанию для стиля 'Normal': {e}. Word будет использовать свои настройки.")

    _add_formatted_paragraph(document, title, font_size=18, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                             font_name=default_font_name)
    _add_formatted_paragraph(document, f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M')}", font_size=10,
                             alignment=WD_ALIGN_PARAGRAPH.CENTER, font_name=default_font_name)
    document.add_paragraph()

    if not grouped_data:
        logger.info("Нет сгруппированных данных для добавления в DOCX.")
        _add_formatted_paragraph(document, "Нет задач для отображения в этом релизе.", font_name=default_font_name)
        try:
            document.save(output_filename)
            logger.info(f"DOCX '{output_filename}' успешно создан (с сообщением 'без задач').")
            return True
        except Exception as e:
            logger.error(f"Ошибка при сохранении пустого DOCX '{output_filename}': {e}", exc_info=True)
            return False

    sorted_ms_versions = sorted(grouped_data.keys())
    logger.info(f"Будет обработано {len(sorted_ms_versions)} версий микросервисов для DOCX.")

    for ms_version in sorted_ms_versions:
        _add_formatted_paragraph(document, ms_version, font_size=16, bold=True, font_name=default_font_name)
        data_for_ms = grouped_data[ms_version]

        current_base_indent = 0.0
        if use_issue_type_grouping and isinstance(data_for_ms, dict):
            sorted_issue_types = sorted(data_for_ms.keys())
            for issue_type in sorted_issue_types:
                _add_formatted_paragraph(document, issue_type, font_size=14, bold=True, left_indent_inches=0.25,
                                         font_name=default_font_name)
                current_base_indent = 0.50  # Отступ для задач под типом
                tasks_to_process = data_for_ms[issue_type]

                for task_num, task in enumerate(tasks_to_process):
                    logger.debug(
                        f"Добавление задачи {task.get('key', 'N/A')} (тип: {issue_type}, версия: {ms_version})")
                    p_key = _add_formatted_paragraph(document, task['key'] + ":", bold=True,
                                                     left_indent_inches=current_base_indent,
                                                     font_name=default_font_name)

                    cust_desc_text = sanitize_text_docx(task['cust_desc'])
                    is_cust_desc_empty = not bool(task['cust_desc'])
                    _add_formatted_paragraph(document,
                                             cust_desc_text if not is_cust_desc_empty else "Описание для клиента отсутствует.",
                                             italic=is_cust_desc_empty,
                                             left_indent_inches=current_base_indent, font_name=default_font_name)

                    if task['install_instr']:
                        _add_formatted_paragraph(document, "Инструкция по установке:", bold=True,
                                                 left_indent_inches=current_base_indent, font_name=default_font_name)
                        _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                                 left_indent_inches=current_base_indent, font_name=default_font_name)

                    if task_num < len(tasks_to_process) - 1:  # Не добавлять отступ после последней задачи в группе
                        document.add_paragraph()
        else:  # Без группировки по типу
            current_base_indent = 0.25  # Отступ для задач под версией микросервиса
            tasks_to_process = data_for_ms
            for task_num, task in enumerate(tasks_to_process):
                logger.debug(f"Добавление задачи {task.get('key', 'N/A')} (версия: {ms_version})")
                # ... (аналогично блоку выше, но с current_base_indent = 0.25) ...
                p_key = _add_formatted_paragraph(document, task['key'] + ":", bold=True,
                                                 left_indent_inches=current_base_indent, font_name=default_font_name)
                cust_desc_text = sanitize_text_docx(task['cust_desc'])
                is_cust_desc_empty = not bool(task['cust_desc'])
                _add_formatted_paragraph(document,
                                         cust_desc_text if not is_cust_desc_empty else "Описание для клиента отсутствует.",
                                         italic=is_cust_desc_empty, left_indent_inches=current_base_indent,
                                         font_name=default_font_name)
                if task['install_instr']:
                    _add_formatted_paragraph(document, "Инструкция по установке:", bold=True,
                                             left_indent_inches=current_base_indent, font_name=default_font_name)
                    _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                             left_indent_inches=current_base_indent, font_name=default_font_name)
                if task_num < len(tasks_to_process) - 1:
                    document.add_paragraph()

        document.add_paragraph()  # Отступ после секции микросервиса

    try:
        document.save(output_filename)
        logger.info(f"DOCX '{output_filename}' успешно создан.")
        return True
    except Exception as e:
        logger.error(f"Ошибка при сохранении DOCX '{output_filename}': {e}", exc_info=True)
        return False