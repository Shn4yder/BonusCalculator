import sys
import os

# Add the parent directory to sys.path to allow imports if run directly
# though running as a module is preferred.
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from bonus_calculator.mpp_parser import load_unique_resources, collect_timephased_data, get_project_name
    from bonus_calculator.excel_utils import load_bonuses_from_excel
    from bonus_calculator.report_generator import generate_report
except ImportError:
    # Fallback for relative imports if run as a module
    from .mpp_parser import load_unique_resources, collect_timephased_data, get_project_name
    from .excel_utils import load_bonuses_from_excel
    from .report_generator import generate_report

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

    # Определяем, является ли report_path путем к файлу или директорией
    final_report_path = report_path
    if report_path.lower().endswith(".xlsx"):
        # Это файл
        parent_dir = os.path.dirname(report_path)
        if parent_dir and not os.path.exists(parent_dir):
            try:
                os.makedirs(parent_dir, exist_ok=True)
            except Exception:
                pass
    else:
        # Это директория
        if not os.path.exists(report_path):
            try:
                os.makedirs(report_path, exist_ok=True)
            except Exception:
                pass
        final_report_path = os.path.join(report_path, "output_data.xlsx")

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

    # Используем все ресурсы по умолчанию
    chosen = resources
    
    if not chosen:
        print("Не выбрано ни одного ресурса.")
        sys.exit(0)

    # 3. Collect Data
    print("Сбор данных о трудозатратах...")
    try:
        project, selected_resources, sorted_months, res_data = collect_timephased_data(mpp_path, chosen)
    except Exception as e:
        print(f"Ошибка при сборе данных: {e}")
        sys.exit(1)
        
    # 4. Get Project Name
    project_name = get_project_name(project)
    
    # 5. Generate Report
    print("Генерация отчета...")
    try:
        generate_report(
            final_report_path,
            project_name,
            sorted_months,
            selected_resources,
            res_data,
            staff_bonus,
            manager_bonus
        )
    except Exception as e:
        print(f"Ошибка при генерации отчета: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
