import os
import shutil
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from .utils import RUSSIAN_MONTHS

def generate_report(
    report_path: str,
    project_name: str,
    sorted_months: list[tuple[int, int]],
    selected_resources: list[tuple[str, object]],
    res_data: dict[int, dict[tuple[int, int], float]],
    staff_bonus: float | None,
    manager_bonus: float | None
):
    # Ensure directory exists
    parent_dir = os.path.dirname(report_path)
    if parent_dir and not os.path.exists(parent_dir):
        try:
            os.makedirs(parent_dir, exist_ok=True)
        except Exception:
            pass

    # Load template or create new
    template_path = os.path.join(os.path.dirname(__file__), "..", "data", "templates", "output_data.xlsx")
    template_path = os.path.abspath(template_path)
    
    wb = None
    if os.path.exists(report_path):
        try:
            wb = openpyxl.load_workbook(report_path)
        except Exception:
            pass
            
    if wb is None:
        if os.path.isfile(template_path):
            try:
                shutil.copyfile(template_path, report_path)
                wb = openpyxl.load_workbook(report_path)
            except Exception:
                wb = openpyxl.Workbook()
        else:
            wb = openpyxl.Workbook()
        
    ws = wb.active

    # Находим строку заголовков (обычно строка 2 или 3, будем искать "ФИО" или "Ресурс")
    header_row = 2
    fio_col = 2  # B
    
    # Поиск заголовков в шаблоне
    found_header = False
    for r_idx in range(1, 10):
        for c_idx in range(1, 10):
            val = ws.cell(row=r_idx, column=c_idx).value
            if val and isinstance(val, str) and ("ФИО" in val or "Ресурс" in val):
                header_row = r_idx
                fio_col = c_idx
                found_header = True
                break
        if found_header:
            break
            
    # Если не нашли, предполагаем стандартную структуру
    if not found_header:
        header_row = 3
        fio_col = 2
    
    # Заполняем месяцы в заголовке
    month_start_col = fio_col + 1
    
    # Очистка области заголовков от объединений и старых данных
    total_cols_needed = len(sorted_months) + 3
    max_col_needed = month_start_col + total_cols_needed
    
    ranges_to_unmerge = []
    for rng in ws.merged_cells.ranges:
        if rng.max_col >= month_start_col:
            ranges_to_unmerge.append(rng)
            
    for rng in list(ranges_to_unmerge):
        try:
            ws.unmerge_cells(str(rng))
        except Exception:
            pass

    # Очищаем две строки заголовков (Год и Месяц)
    start_header_clear = header_row - 1 if header_row > 1 else header_row
    for r in range(start_header_clear, header_row + 1):
        for c in range(month_start_col, max_col_needed + 5):
            ws.cell(row=r, column=c).value = None
    
    # Заполняем месяцы и годы
    current_year_start_col = None
    current_year = None
    
    for i, (y, m) in enumerate(sorted_months):
        col_idx = month_start_col + i
        
        # Год (в строке над месяцем)
        if header_row > 1:
            y_cell = ws.cell(row=header_row-1, column=col_idx)
            y_cell.value = y
            y_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Логика объединения годов
            if current_year is None:
                current_year = y
                current_year_start_col = col_idx
            elif y != current_year:
                # Объединяем предыдущий год
                if current_year_start_col < col_idx - 1:
                    ws.merge_cells(start_row=header_row-1, start_column=current_year_start_col, 
                                   end_row=header_row-1, end_column=col_idx-1)
                current_year = y
                current_year_start_col = col_idx
        
        # Месяц
        month_name = RUSSIAN_MONTHS.get(m, "")
        c = ws.cell(row=header_row, column=col_idx)
        c.value = month_name
        c.alignment = Alignment(horizontal='center', vertical='center')

    # Объединяем последний год
    if header_row > 1 and current_year is not None:
        last_col = month_start_col + len(sorted_months) - 1
        if current_year_start_col < last_col:
             ws.merge_cells(start_row=header_row-1, start_column=current_year_start_col, 
                            end_row=header_row-1, end_column=last_col)

    # Столбец "Итого"
    total_work_col = month_start_col + len(sorted_months)
    
    # Столбец "КТУ"
    ktu_col = total_work_col + 1
    
    # Столбец "Премия"
    bonus_col = ktu_col + 1
    
    # Заголовки Итого/КТУ/Премия
    if header_row > 1:
        # Общий итог (остается под заголовком проекта, занимает 2 строки: Год и Месяц)
        ws.cell(row=header_row-1, column=total_work_col, value="Общий итог")
        ws.merge_cells(start_row=header_row-1, start_column=total_work_col, end_row=header_row, end_column=total_work_col)
        c_tot = ws.cell(row=header_row-1, column=total_work_col)
        c_tot.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c_tot.font = Font(bold=False, italic=False)
        
        # КТУ (объединяется с 1-й строки до строки заголовков)
        ws.cell(row=1, column=ktu_col, value="КТУ")
        ws.merge_cells(start_row=1, start_column=ktu_col, end_row=header_row, end_column=ktu_col)
        c_ktu_h = ws.cell(row=1, column=ktu_col)
        c_ktu_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c_ktu_h.font = Font(bold=False, italic=False)
        
        # Премия (объединяется с 1-й строки до строки заголовков)
        ws.cell(row=1, column=bonus_col, value="Премия")
        ws.merge_cells(start_row=1, start_column=bonus_col, end_row=header_row, end_column=bonus_col)
        c_bon_h = ws.cell(row=1, column=bonus_col)
        c_bon_h.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c_bon_h.font = Font(bold=True)
    else:
        ws.cell(row=header_row, column=total_work_col, value="Общий итог")
        ws.cell(row=header_row, column=ktu_col, value="КТУ")
        ws.cell(row=header_row, column=bonus_col, value="Премия")

    # -- MERGING A1 --
    ws["A1"] = f'Расчет премии исполнителей по проекту\n"{project_name}"'
         
    try:
        # Объединяем A1 до столбца ПЕРЕД КТУ (то есть до total_work_col включительно)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_work_col)
        ws["A1"].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws["A1"].font = Font(bold=True)
    except Exception:
        pass

    # Заполняем данные
    start_data_row = header_row + 1
    
    # Очистка строк данных
    for r in range(start_data_row, start_data_row + 100):
        if ws.cell(row=r, column=fio_col).value is None:
            break
        
    for idx, (name, r_obj) in enumerate(selected_resources):
        row = start_data_row + idx
        ws.cell(row=row, column=1, value=idx + 1)
        ws.cell(row=row, column=fio_col, value=name)
        
        monthly_data = res_data.get(idx, {})
        
        # Заполняем часы
        for i, (y, m) in enumerate(sorted_months):
            col_idx = month_start_col + i
            hours = monthly_data.get((y, m), 0.0)
            c_h = ws.cell(row=row, column=col_idx, value=hours)
            c_h.number_format = '0.00'
            c_h.alignment = Alignment(horizontal='center', vertical='center')
            c_h.font = Font(bold=False, italic=False)
            
        # Формула Итого
        start_col_letter = openpyxl.utils.get_column_letter(month_start_col)
        end_col_letter = openpyxl.utils.get_column_letter(total_work_col - 1)
        c_sum = ws.cell(row=row, column=total_work_col, value=f"=ROUND(SUM({start_col_letter}{row}:{end_col_letter}{row}), 2)")
        c_sum.number_format = '0.00'
        c_sum.alignment = Alignment(horizontal='center', vertical='center')
        c_sum.font = Font(bold=False, italic=False)

    # Строка ИТОГО
    total_row_idx = start_data_row + len(selected_resources)
    c_total_label = ws.cell(row=total_row_idx, column=fio_col, value="Общий итог")
    c_total_label.font = Font(bold=False, italic=False)
    c_total_label.alignment = Alignment(horizontal='center', vertical='center')
    
    # Суммы по месяцам
    for i in range(len(sorted_months)):
        col_idx = month_start_col + i
        col_let = openpyxl.utils.get_column_letter(col_idx)
        c_tot_m = ws.cell(row=total_row_idx, column=col_idx, value=f"=ROUND(SUM({col_let}{start_data_row}:{col_let}{total_row_idx-1}), 2)")
        c_tot_m.number_format = '0.00'
        c_tot_m.alignment = Alignment(horizontal='center', vertical='center')
        c_tot_m.font = Font(bold=False, italic=False)
        
    # Сумма общих итогов
    total_work_col_letter = openpyxl.utils.get_column_letter(total_work_col)
    c_tot_w = ws.cell(row=total_row_idx, column=total_work_col, value=f"=ROUND(SUM({total_work_col_letter}{start_data_row}:{total_work_col_letter}{total_row_idx-1}), 2)")
    c_tot_w.number_format = '0.00'
    c_tot_w.alignment = Alignment(horizontal='center', vertical='center')
    c_tot_w.font = Font(bold=False, italic=False)
    
    # Сумма КТУ
    ktu_col_letter = openpyxl.utils.get_column_letter(ktu_col)
    c_ktu_tot = ws.cell(row=total_row_idx, column=ktu_col, value=f"=ROUND(SUM({ktu_col_letter}{start_data_row}:{ktu_col_letter}{total_row_idx-1}), 2)")
    c_ktu_tot.number_format = '0.00'
    c_ktu_tot.alignment = Alignment(horizontal='center', vertical='center')
    c_ktu_tot.font = Font(bold=False, italic=False)
    
    # Сумма Премий
    bonus_col_letter = openpyxl.utils.get_column_letter(bonus_col)
    c_bon_tot = ws.cell(row=total_row_idx, column=bonus_col, value=f"=ROUND(SUM({bonus_col_letter}{start_data_row}:{bonus_col_letter}{total_row_idx-1}), 2)")
    c_bon_tot.number_format = '0.00'
    c_bon_tot.alignment = Alignment(horizontal='center', vertical='center')
    c_bon_tot.font = Font(bold=True)

    # КТУ и Премии
    
    # 1. Сначала вычислим общее количество часов, чтобы найти того, у кого их больше всего
    total_hours_per_resource = {}
    for idx in range(len(selected_resources)):
        monthly_data = res_data.get(idx, {})
        total_hours_per_resource[idx] = sum(monthly_data.values())
        
    max_work_idx = -1
    max_val = -1.0
    
    if len(selected_resources) > 0:
        max_work_idx = 0
        max_val = total_hours_per_resource[0]
        for idx, val in total_hours_per_resource.items():
            if val > max_val:
                max_val = val
                max_work_idx = idx

    for idx in range(len(selected_resources)):
        row = start_data_row + idx
        
        # КТУ
        if idx == max_work_idx and len(selected_resources) > 1:
            # Plug formula: 100 - SUM(others)
            parts = []
            if idx > 0:
                prev_start = start_data_row
                prev_end = row - 1
                parts.append(f"{ktu_col_letter}{prev_start}:{ktu_col_letter}{prev_end}")
            
            if idx < len(selected_resources) - 1:
                next_start = row + 1
                next_end = start_data_row + len(selected_resources) - 1
                parts.append(f"{ktu_col_letter}{next_start}:{ktu_col_letter}{next_end}")
                
            sum_formula = "+".join([f"SUM({p})" for p in parts])
            c_ktu = ws.cell(row=row, column=ktu_col, value=f"=100-({sum_formula})")
        else:
            # Standard formula: =ROUND(TotalWork * 100 / GrandTotalWork, 2)
            c_ktu = ws.cell(row=row, column=ktu_col, value=f"=ROUND({total_work_col_letter}{row}*100/${total_work_col_letter}${total_row_idx}, 2)")
            
        c_ktu.number_format = '0.00'
        c_ktu.alignment = Alignment(horizontal='center', vertical='center')
        c_ktu.font = Font(bold=False, italic=False)
        
        # Премия
        if staff_bonus is not None:
            if idx == max_work_idx and len(selected_resources) > 1:
                # Plug formula: StaffBonus - SUM(others)
                parts = []
                if idx > 0:
                    prev_start = start_data_row
                    prev_end = row - 1
                    parts.append(f"{bonus_col_letter}{prev_start}:{bonus_col_letter}{prev_end}")
                
                if idx < len(selected_resources) - 1:
                    next_start = row + 1
                    next_end = start_data_row + len(selected_resources) - 1
                    parts.append(f"{bonus_col_letter}{next_start}:{bonus_col_letter}{next_end}")
                    
                sum_formula = "+".join([f"SUM({p})" for p in parts])
                c_bon = ws.cell(row=row, column=bonus_col, value=f"={staff_bonus}-({sum_formula})")
            else:
                # Standard formula: =ROUND(KTU / 100 * StaffBonus, 2)
                c_bon = ws.cell(row=row, column=bonus_col, value=f"=ROUND({ktu_col_letter}{row}/100*{staff_bonus}, 2)")
        else:
             c_bon = ws.cell(row=row, column=bonus_col, value=0)
             
        c_bon.number_format = '0.00'
        c_bon.alignment = Alignment(horizontal='center', vertical='center')
        c_bon.font = Font(bold=True)

    # Строка "Премия руководителя"
    manager_row = total_row_idx + 2
    ws.cell(row=manager_row, column=fio_col, value="Премия РП")
    c_m = ws.cell(row=manager_row, column=bonus_col, value=manager_bonus if manager_bonus is not None else 0)
    c_m.number_format = '0.00'
    c_m.alignment = Alignment(horizontal='center', vertical='center')
    c_m.font = Font(bold=True)

    # -- BORDERS --
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
                         
    # Main table
    start_border_row = header_row - 1 if header_row > 1 else header_row
    end_border_row = total_row_idx
    
    for r in range(start_border_row, end_border_row + 1):
        for c in range(1, bonus_col + 1):
            ws.cell(row=r, column=c).border = thin_border
            
    # Manager row - NO BORDERS
    # Explicitly removing borders for manager row (just in case they were inherited)
    ws.cell(row=manager_row, column=fio_col).border = None
    ws.cell(row=manager_row, column=bonus_col).border = None

    # -- COLUMN A FORMATTING --
    # "все данные в столбце а кроме ячейки а1 должны быть расположены по левому краю и иметь шрифт calibri 
    # (кроме ячеек с надписью премия рп, заказчик/куратор, руководитель, ивестор)"
    
    # Определяем ключевые слова для исключения
    exclude_keywords = [
        "премия рп", 
        "заказчик", 
        "куратор", 
        "руководитель", 
        "инвестор"
    ]
    
    calibri_font = Font(name='Calibri', size=11, bold=False)
    left_align = Alignment(horizontal='left', vertical='center')
    
    # Проходим по всем заполненным строкам столбца A, начиная с A2
    max_row = ws.max_row
    for r in range(2, max_row + 1):
        cell_a = ws.cell(row=r, column=1)
        val_a = str(cell_a.value).strip().lower() if cell_a.value else ""
        
        # Проверяем, не содержит ли ячейка запрещенные слова
        should_skip = False
        for kw in exclude_keywords:
            if kw in val_a:
                should_skip = True
                break
        
        # Также проверяем соседнюю ячейку (столбец B), так как "Премия РП" пишется в fio_col (B)
        # Если в строке есть "Премия РП" в столбце B, возможно, пользователь хочет пропустить форматирование A в этой строке?
        # По запросу: "кроме ячеек с надписью..." - скорее всего, речь о самой ячейке.
        # Но "Премия РП" в коде пишется в fio_col.
        # Если "Премия РП" в B, а A пустое -> форматирование пустого A не повредит.
        # Если "Руководитель проекта" (подпись) в A -> тогда skip.
        
        if not should_skip:
             cell_a.font = calibri_font
             cell_a.alignment = left_align

    # Сохраняем
    try:
        wb.save(report_path)
        print(f"Отчет успешно сформирован: {report_path}")
    except Exception as e:
        print(f"Ошибка при сохранении отчета: {e}")
