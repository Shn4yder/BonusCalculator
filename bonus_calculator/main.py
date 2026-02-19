import sys
import os

# Add the parent directory to sys.path to allow imports if run directly
# though running as a module is preferred.
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from bonus_calculator.mpp_parser import load_unique_resources, collect_timephased_data, get_project_name, is_project_completed
    from bonus_calculator.excel_utils import load_bonuses_from_excel
    from bonus_calculator.report_generator import generate_report
    from bonus_calculator.utils import parse_indices, get_unique_report_path
except ImportError:
    # Fallback for relative imports if run as a module
    from .mpp_parser import load_unique_resources, collect_timephased_data, get_project_name, is_project_completed
    from .excel_utils import load_bonuses_from_excel
    from .report_generator import generate_report
    from .utils import parse_indices, get_unique_report_path

def resolve_file(p: str) -> str:
    p1 = os.path.abspath(p)
    if os.path.isfile(p1):
        return p1
    base = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    p2 = os.path.abspath(os.path.join(base, p))
    if os.path.isfile(p2):
        return p2
    return p1

def main():
    if len(sys.argv) != 4:
        print("Использование: python -m bonus_calculator.main <путь_к_файлу.mpp> <путь_к_файлу.xlsx> <путь_для_отчета>")
        sys.exit(2)
        
    mpp_arg = sys.argv[1]
    xlsx_arg = sys.argv[2]
    report_path = sys.argv[3]
    
    mpp_path = resolve_file(mpp_arg)
    xlsx_path = resolve_file(xlsx_arg)
    
    if not os.path.isfile(mpp_path) or not mpp_path.lower().endswith(".mpp"):
        print("Ошибка: первый аргумент должен быть существующим файлом с расширением .mpp")
        sys.exit(2)
        
    if not os.path.isfile(xlsx_path) or not xlsx_path.lower().endswith(".xlsx"):
        print("Ошибка: второй аргумент должен быть существующим файлом с расширением .xlsx")
        sys.exit(2)

    # Check project completion status
    print("Проверка статуса проекта...")
    try:
        if not is_project_completed(mpp_path):
            print("Внимание: Проект не завершен на 100%. Отчет не будет сформирован.")
            sys.exit(0)
    except Exception as e:
        print(f"Ошибка при проверке статуса проекта: {e}")
        sys.exit(1)

    # Определяем, является ли report_path путем к файлу или директорией
    # final_report_path logic moved to after project name retrieval
    pass

    # 1. Load Bonuses
    staff_bonus = None
    manager_bonus = None
    try:
        staff_bonus, manager_bonus = load_bonuses_from_excel(xlsx_path)
        if staff_bonus is not None:
            staff_bonus = round(staff_bonus, 2)
        if manager_bonus is not None:
            manager_bonus = round(manager_bonus, 2)
            
        print("Премии из Excel:")
        print(f"Премия исполнителей (общий итог): {staff_bonus}")
        print(f"Премия руководителя (общий итог): {manager_bonus}")
    except Exception as e:
        print(f"Ошибка чтения премий из Excel: {e}")

    # 2. Load Resources
    try:
        resources = load_unique_resources(mpp_path)
    except Exception as e:
        print(f"Ошибка при чтении .mpp: {e}")
        sys.exit(1)
        
    if not resources:
        print("Ресурсы не найдены.")
        sys.exit(0)
        
    for i, r in enumerate(resources, start=1):
        print(f"{i}. {r}")

    # Запрашиваем выбор ресурсов у пользователя
    print("\nВведите номера ресурсов, которые необходимо включить в отчет (через запятую или пробел).")
    print("Если ничего не ввести и нажать Enter, будут выбраны все ресурсы.")
    selection = input("Ваш выбор: ")
    
    chosen = []
    if not selection.strip():
        chosen = resources
    else:
        indices = parse_indices(selection, len(resources))
        chosen = [resources[i - 1] for i in indices]
    
    if not chosen:
        print("Не выбрано ни одного ресурса.")
        sys.exit(0)

    # 3. Collect Data
    print("Сбор данных о трудозатратах (для всех ресурсов)...")
    try:
        # Collect data for ALL resources to ensure correct totals
        project, all_resources_data, sorted_months, res_data = collect_timephased_data(mpp_path, resources)
    except Exception as e:
        print(f"Ошибка при сборе данных: {e}")
        sys.exit(1)
        
    # Determine visible indices
    # We need to match chosen names to the returned all_resources_data
    # all_resources_data is a list of (name, obj)
    
    visible_indices = []
    chosen_set = set(chosen)
    
    for idx, (name, r_obj) in enumerate(all_resources_data):
        if name in chosen_set:
            visible_indices.append(idx)
            
    if not visible_indices:
         print("Внимание: выбранные ресурсы не найдены в данных проекта.")
        
    # 4. Get Project Name
    project_name = get_project_name(project)
    
    # 4.5 Determine Final Report Path
    final_report_path = get_unique_report_path(report_path, project_name)
    print(f"Отчет будет сохранен как: {final_report_path}")
    
    # 5. Generate Report
    print("Генерация отчета...")
    try:
        generate_report(
            final_report_path,
            project_name,
            sorted_months,
            all_resources_data, # Pass ALL resources
            res_data,           # Pass ALL data
            staff_bonus,
            manager_bonus,
            visible_indices=visible_indices # Pass filter
        )
    except Exception as e:
        print(f"Ошибка при генерации отчета: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
