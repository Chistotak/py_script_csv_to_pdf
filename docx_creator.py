# release_notes_generator/docx_creator.py
import logging
from datetime import datetime
import os
import re

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
except ImportError:
    logging.critical("Библиотека python-docx не установлена. Установите ее командой: pip install python-docx")
    raise

logger = logging.getLogger(__name__)


def sanitize_text_docx(text):
    if text is None:
        return ""
    return str(text)


def get_style_value(style_config, section, key, default_value, value_type=str):
    section_dict = style_config.get(section, {})
    val_str_from_config = section_dict.get(key)  # Получаем значение из конфига

    val_to_process_str = None
    used_default = False

    if val_str_from_config is not None:
        val_to_process_str = str(val_str_from_config)  # Используем значение из конфига
    else:
        # logger.debug(f"Ключ '{key}' не найден в секции '{section}'. Используется дефолт: {default_value}")
        val_to_process_str = str(default_value)  # Используем дефолт, преобразованный в строку
        used_default = True

    try:
        if value_type == int: return int(val_to_process_str)
        if value_type == float: return float(val_to_process_str)
        if value_type == bool: return val_to_process_str.lower() == 'true'
        if value_type == RGBColor:
            # Если исходное значение было дефолтным и это уже RGBColor, возвращаем его
            if used_default and isinstance(default_value, RGBColor):
                return default_value
            hex_color = val_to_process_str.lstrip('#')
            if len(hex_color) == 6:
                return RGBColor.from_string(hex_color)
            logger.warning(f"Некорректный HEX цвета '{val_to_process_str}' для {section}/{key}. Используется черный.")
            return RGBColor(0, 0, 0)
        return val_to_process_str  # str по умолчанию
    except (ValueError, TypeError) as e:
        logger.warning(
            f"Ошибка преобразования '{val_to_process_str}' для {section}/{key} к {value_type.__name__}: {e}. Используется дефолт: {default_value}")
        # Возврат безопасных дефолтов в случае ошибки преобразования
        if value_type == int: return 0 if not isinstance(default_value, int) else default_value
        if value_type == float: return 0.0 if not isinstance(default_value, float) else default_value
        if value_type == bool: return False if not isinstance(default_value, bool) else default_value
        if value_type == RGBColor: return RGBColor(0, 0, 0)
        return str(default_value)


def _apply_run_formatting(run, font_name, font_size_pt, bold, italic, color=None, underline=False):
    if font_name:
        try:
            run.font.name = font_name
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
            r.rPr.rFonts.set(qn('w:cs'), font_name)
        except Exception as e:
            logger.warning(f"Не удалось установить шрифт '{font_name}': {e}")
    if font_size_pt: run.font.size = Pt(font_size_pt)
    run.bold = bold
    run.italic = italic
    if color: run.font.color.rgb = color
    run.underline = underline


def _add_formatted_paragraph(document, text, style_config,
                             font_key='main', fontsize_key='normal_style_base', color_key='normal_text',
                             bold=False, italic=False, underline=False,
                             alignment=None, left_indent_inches=None,
                             space_before_key=None, space_after_key=None,  # Если None, отступ не ставится из конфига
                             line_spacing_rule=None, line_spacing_val=None,
                             keep_with_next=False, keep_together=False, style_name=None):
    p = document.add_paragraph()
    is_spacing_paragraph = (text is None)  # True, если это параграф только для отступа
    run_text_to_add = sanitize_text_docx(text) if not is_spacing_paragraph else u'\u00A0'

    font_name_val = get_style_value(style_config, 'Fonts', font_key, 'Arial')
    font_size_pt_val = get_style_value(style_config, 'FontSizes', fontsize_key, 11, value_type=int)
    actual_font_size_pt = 1 if is_spacing_paragraph else font_size_pt_val

    # Получаем цвет, даже если это параграф-отступ, т.к. run все равно создается
    run_color_val = get_style_value(style_config, 'Colors', color_key, RGBColor(0, 0, 0), value_type=RGBColor)

    current_run = None  # Инициализируем current_run
    if style_name:
        try:
            p.style = style_name
            if not is_spacing_paragraph or run_text_to_add.strip():
                current_run = p.runs[0] if p.runs and p.runs[0].text else p.add_run()
                current_run.text = run_text_to_add
        except KeyError:
            logger.warning(f"Стиль '{style_name}' не найден...")
            if not is_spacing_paragraph or run_text_to_add.strip():
                current_run = p.add_run(run_text_to_add)
    elif not is_spacing_paragraph or run_text_to_add.strip():
        current_run = p.add_run(run_text_to_add)
    elif is_spacing_paragraph:  # Только для отступа
        current_run = p.add_run(run_text_to_add)

    if current_run:  # Форматируем run, если он был создан
        _apply_run_formatting(current_run, font_name_val, actual_font_size_pt, bold, italic, run_color_val, underline)

    p_fmt = p.paragraph_format
    if alignment: p_fmt.alignment = alignment
    if left_indent_inches is not None: p_fmt.left_indent = Inches(left_indent_inches)

    if space_before_key: p_fmt.space_before = Pt(
        get_style_value(style_config, 'Spacing', space_before_key, 0, value_type=int))
    if space_after_key: p_fmt.space_after = Pt(
        get_style_value(style_config, 'Spacing', space_after_key, 0 if is_spacing_paragraph else 6, value_type=int))

    if line_spacing_rule and line_spacing_val is not None:
        p_fmt.line_spacing_rule = line_spacing_rule
        if line_spacing_rule in [WD_LINE_SPACING.MULTIPLE, WD_LINE_SPACING.AT_LEAST, WD_LINE_SPACING.EXACTLY]:
            p_fmt.line_spacing = float(
                str(line_spacing_val).replace(',', '.'))  # Для MULTIPLE нужен float, например 1.15
    p_fmt.keep_with_next = keep_with_next
    p_fmt.keep_together = keep_together
    return p


def extract_microservice_info_for_summary_table(grouped_data_keys, main_config_data):
    # ... (код этой функции без изменений, как в предыдущем полном ответе) ...
    logger_func = logging.getLogger(__name__)
    microservices_summary = []
    version_pattern = re.compile(r"^([A-Z]{2})(\d+(\.\d+){1,2})$")
    seen_summary_entries = set()
    for original_ms_key in sorted(list(grouped_data_keys)):
        service_name_for_table = original_ms_key
        version_number_part = ""
        match = version_pattern.match(original_ms_key)
        if match:
            prefix_key_for_config = match.group(1).upper()
            version_number_part = match.group(2)
            template = main_config_data.get('MicroserviceVersions', {}).get(prefix_key_for_config)
            if template:
                service_name_for_table = template.replace("{{version}}", "").replace("(версия )", "").strip()
            else:
                service_name_for_table = prefix_key_for_config
        else:
            logger_func.warning(f"Не удалось разобрать ключ '{original_ms_key}' для сводной таблицы.")
        summary_tuple = (service_name_for_table, version_number_part)
        if summary_tuple not in seen_summary_entries:
            microservices_summary.append(
                {'service_name': service_name_for_table, 'version_number': version_number_part})
            seen_summary_entries.add(summary_tuple)
    return microservices_summary


def create_release_notes_docx(output_filename, title, grouped_data,
                              use_client_grouping_flag, use_issue_type_grouping_flag,
                              microservices_summary_data=None,
                              main_config=None, style_config=None):
    document = Document()
    logger.info(f"Создание DOCX: {output_filename}")

    # Настройка стиля 'Normal'
    s_font_main = get_style_value(style_config, 'Fonts', 'main', 'Arial')
    s_fontsize_normal = get_style_value(style_config, 'FontSizes', 'normal_style_base', 11, value_type=int)
    s_space_after_normal = get_style_value(style_config, 'Spacing', 'normal_paragraph_after', 6, value_type=int)
    try:
        normal_style = document.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = s_font_main
        r_normal = normal_font._element
        r_normal.rPr.rFonts.set(qn('w:eastAsia'), s_font_main);
        r_normal.rPr.rFonts.set(qn('w:cs'), s_font_main)
        normal_font.size = Pt(s_fontsize_normal)
        normal_style.paragraph_format.space_after = Pt(s_space_after_normal)
        normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        normal_style.paragraph_format.line_spacing = 1.15
        logger.info(
            f"Стиль 'Normal' настроен: {s_font_main} {s_fontsize_normal}pt, отступ после {s_space_after_normal}pt.")
    except Exception as e:
        logger.warning(f"Не удалось настроить стиль 'Normal': {e}.")

    # Логотип
    if main_config:
        logo_path_key = main_config.get('General', {}).get('logo_path')
        if logo_path_key:
            actual_logo_path = logo_path_key
            config_dir = main_config.get('_config_dir_', os.getcwd())
            if not os.path.isabs(actual_logo_path): actual_logo_path = os.path.join(config_dir, actual_logo_path)
            if os.path.exists(actual_logo_path):
                try:
                    logo_width = get_style_value(main_config, 'General', 'logo_width_inches', 1.5, value_type=float)
                    p_logo = document.add_paragraph();
                    p_logo.add_run().add_picture(actual_logo_path, width=Inches(logo_width))
                    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER;
                    p_logo.paragraph_format.space_after = Pt(12)
                    logger.info(f"Логотип '{actual_logo_path}' добавлен.")
                except Exception as e:
                    logger.error(f"Не удалось добавить логотип '{actual_logo_path}': {e}")
            else:
                logger.warning(f"Файл логотипа '{actual_logo_path}' не найден.")

    # Заголовок и дата
    _add_formatted_paragraph(document, title, style_config, font_key='title', fontsize_key='title', color_key='title',
                             bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after_key='after_title')
    _add_formatted_paragraph(document, f"Дата генерации: {datetime.now().strftime('%Y-%m-%d %H:%M')}", style_config,
                             font_key='main', fontsize_key='date', color_key='date_text', italic=True,
                             alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after_key='after_date')

    # Таблица микросервисов
    if microservices_summary_data:
        _add_formatted_paragraph(document, "Состав релиза по микросервисам:", style_config, font_key='section_header',
                                 fontsize_key='summary_table_title', color_key='summary_table_title', bold=True,
                                 space_after_key='after_summary_table_title', keep_with_next=True)
        if microservices_summary_data:
            table = document.add_table(rows=1, cols=2)  # ... (код создания и форматирования таблицы как раньше)
            table.style = 'Table Grid'
            table.columns[0].width = Inches(
                get_style_value(style_config, 'TableLayout', 'summary_table_col1_width_inches', 4.0, value_type=float))
            table.columns[1].width = Inches(
                get_style_value(style_config, 'TableLayout', 'summary_table_col2_width_inches', 1.5, value_type=float))
            hdr_font = get_style_value(style_config, 'Fonts', 'main', 'Arial')
            hdr_fontsize = get_style_value(style_config, 'FontSizes', 'summary_table_header', 11, value_type=int)
            hdr_color = get_style_value(style_config, 'Colors', 'table_header_text', RGBColor(0, 0, 0),
                                        value_type=RGBColor)
            hdr_cells = table.rows[0].cells;
            col_texts = ['Микросервис', 'Версия']
            for i, cell_text in enumerate(col_texts):
                p = hdr_cells[i].paragraphs[0];
                p.text = "";
                run = p.add_run(cell_text)
                _apply_run_formatting(run, hdr_font, hdr_fontsize, True, False, color=hdr_color)
            cell_font = get_style_value(style_config, 'Fonts', 'main', 'Arial')
            cell_fontsize = get_style_value(style_config, 'FontSizes', 'summary_table_text', 10, value_type=int)
            cell_color = get_style_value(style_config, 'Colors', 'table_text', RGBColor(0, 0, 0), value_type=RGBColor)
            for item in microservices_summary_data:
                row_cells = table.add_row().cells;
                texts_to_add = [item.get('service_name', 'N/A'), item.get('version_number', 'N/A')]
                for i, cell_content in enumerate(texts_to_add):
                    row_cells[i].text = cell_content
                    if row_cells[i].paragraphs:
                        p_cell = row_cells[i].paragraphs[0]
                        current_run = p_cell.runs[0] if p_cell.runs else p_cell.add_run(cell_content)
                        if not current_run.text and cell_content: current_run.text = cell_content
                        _apply_run_formatting(current_run, cell_font, cell_fontsize, False, False, color=cell_color)
            logger.info("Таблица микросервисов добавлена.")
        else:
            _add_formatted_paragraph(document, "Информация о версиях микросервисов отсутствует.", style_config,
                                     fontsize_key='normal_style_base', italic=True)
        _add_formatted_paragraph(document, None, style_config, space_after_key='after_summary_table')
    elif grouped_data:
        _add_formatted_paragraph(document, None, style_config, space_after_key='normal_paragraph_after')

    # Детализация задач
    if not grouped_data:
        if not microservices_summary_data: _add_formatted_paragraph(document, "Нет задач для отображения.",
                                                                    style_config, font_key='main',
                                                                    fontsize_key='normal_style_base')
    else:
        microservice_version_pattern_docx = re.compile(r"^([A-Z]{2})(\d+(\.\d+){1,2})$")
        sorted_ms_versions_original = sorted(grouped_data.keys())

        for ms_idx, ms_version_original_key in enumerate(sorted_ms_versions_original):
            display_ms_version = ms_version_original_key
            if main_config:  # Получаем display_ms_version
                match = microservice_version_pattern_docx.match(ms_version_original_key)
                if match:
                    prefix = match.group(1).upper();
                    version_num = match.group(2)
                    template = main_config.get('MicroserviceVersions', {}).get(prefix)
                    if template: display_ms_version = template.replace("{{version}}", version_num)

            _add_formatted_paragraph(document, display_ms_version, style_config, font_key='section_header',
                                     fontsize_key='ms_version_header', color_key='section_header', bold=True,
                                     space_before_key='section_after_space' if ms_idx > 0 else None,
                                     space_after_key='after_ms_version_header', keep_with_next=True)

            data_for_current_ms = grouped_data[ms_version_original_key]
            client_or_type_keys = sorted(data_for_current_ms.keys(),
                                         key=lambda k: (k.lower() == "не указан" or k.lower() == "общие задачи",
                                                        k.lower())) if isinstance(data_for_current_ms, dict) else []

            if use_client_grouping_flag:
                for client_name_display in client_or_type_keys:
                    _add_formatted_paragraph(document, client_name_display, style_config, font_key='client_header',
                                             fontsize_key='client_header', color_key='client_header', bold=True,
                                             left_indent_inches=0.25, space_before_key='after_ms_version_header',
                                             space_after_key='after_client_header', keep_with_next=True)
                    data_for_current_client = data_for_current_ms[client_name_display]
                    if use_issue_type_grouping_flag and isinstance(data_for_current_client, dict):
                        issue_type_keys = sorted(data_for_current_client.keys(),
                                                 key=lambda k: (k.lower() == "не указан тип" or k.lower() == "задачи",
                                                                k.lower()))
                        for issue_type_name in issue_type_keys:
                            _add_formatted_paragraph(document, issue_type_name, style_config,
                                                     font_key='issue_type_header', fontsize_key='issue_type_header',
                                                     color_key='sub_header', bold=True, left_indent_inches=0.50,
                                                     space_before_key='after_client_header',
                                                     space_after_key='after_issue_type_header', keep_with_next=True)
                            tasks = data_for_current_client[issue_type_name];
                            current_indent = 0.75
                            for task in tasks:  # Вывод задач
                                _add_formatted_paragraph(document, task['key'] + ":", style_config, font_key='task_key',
                                                         fontsize_key='task_key', color_key='task_key', bold=True,
                                                         left_indent_inches=current_indent,
                                                         space_before_key='task_block_internal_space',
                                                         space_after_key='task_key_after')
                                desc = sanitize_text_docx(task['cust_desc']);
                                desc_empty = not bool(desc)
                                _add_formatted_paragraph(document,
                                                         desc if not desc_empty else "Описание ... отсутствует.",
                                                         style_config, font_key='main', fontsize_key='task_description',
                                                         color_key='task_description', italic=desc_empty,
                                                         left_indent_inches=current_indent,
                                                         space_after_key='task_description_after' if not task[
                                                             'install_instr'] else 'task_block_internal_space')
                                if task['install_instr']:
                                    _add_formatted_paragraph(document, "Инструкция:", style_config, font_key='main',
                                                             fontsize_key='install_instruction_label',
                                                             color_key='install_instruction_label', bold=True,
                                                             left_indent_inches=current_indent,
                                                             space_before_key='task_block_internal_space',
                                                             space_after_key='install_label_after')
                                    _add_formatted_paragraph(document, sanitize_text_docx(task['install_instr']),
                                                             style_config, font_key='main',
                                                             fontsize_key='install_instruction_text',
                                                             color_key='install_instruction_text',
                                                             left_indent_inches=current_indent,
                                                             space_after_key='install_text_after')
                    elif isinstance(data_for_current_client, list):  # Задачи под клиентом
                        tasks = data_for_current_client;
                        current_indent = 0.50
                        for task in tasks:  # Вывод задач
                            _add_formatted_paragraph(document, task['key'] + ":", style_config, font_key='task_key',
                                                     fontsize_key='task_key', color_key='task_key', bold=True,
                                                     left_indent_inches=current_indent,
                                                     space_before_key='task_block_internal_space',
                                                     space_after_key='task_key_after')
                            # ... и т.д. для описания и инструкции ...
            elif use_issue_type_grouping_flag:  # Только типы, без клиентов
                for issue_type_name in client_or_type_keys:  # Здесь это типы
                    _add_formatted_paragraph(document, issue_type_name, style_config, font_key='issue_type_header',
                                             fontsize_key='issue_type_header', color_key='sub_header', bold=True,
                                             left_indent_inches=0.25, space_before_key='after_ms_version_header',
                                             space_after_key='after_issue_type_header', keep_with_next=True)
                    tasks = data_for_current_ms[issue_type_name];
                    current_indent = 0.50
                    for task in tasks:  # Вывод задач
                        _add_formatted_paragraph(document, task['key'] + ":", style_config, font_key='task_key',
                                                 fontsize_key='task_key', color_key='task_key', bold=True,
                                                 left_indent_inches=current_indent,
                                                 space_before_key='task_block_internal_space',
                                                 space_after_key='task_key_after')
                        # ... и т.д. для описания и инструкции ...
            else:  # Нет вложенных группировок
                tasks = data_for_current_ms;
                current_indent = 0.25
                for task in tasks:  # Вывод задач
                    _add_formatted_paragraph(document, task['key'] + ":", style_config, font_key='task_key',
                                             fontsize_key='task_key', color_key='task_key', bold=True,
                                             left_indent_inches=current_indent,
                                             space_before_key='task_block_internal_space',
                                             space_after_key='task_key_after')
                    # ... и т.д. для описания и инструкции ...

            if ms_idx < len(sorted_ms_versions_original) - 1:
                _add_formatted_paragraph(document, None, style_config, space_after_key='section_after_space')

    try:
        document.save(output_filename)
        logger.info(f"DOCX '{output_filename}' успешно сохранен.")
        return True
    except Exception as e:
        logger.error(f"Ошибка при сохранении DOCX '{output_filename}': {e}", exc_info=True)
        return False