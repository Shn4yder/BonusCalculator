import sys
import os
import re
from collections import OrderedDict
try:
    import aspose.tasks as tsk
except Exception as e:
    print("Не удалось импортировать библиотеку aspose-tasks. Установите пакет: pip install aspose-tasks")
    sys.exit(1)
try:
    import openpyxl
except Exception:
    openpyxl = None


def normalize_resource_name(name: str) -> str:
    if name is None:
        return ""
    s = str(name).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s.strip())
    return s.casefold()


def get_resource_name(resource) -> str:
    name = None
    if hasattr(resource, "name"):
        try:
            name = resource.name
        except Exception:
            name = None
    if not name:
        try:
            if hasattr(tsk, "Rsc") and hasattr(resource, "get"):
                name = resource.get(tsk.Rsc.NAME)
        except Exception:
            name = None
    return name if isinstance(name, str) else (str(name) if name is not None else "")


def load_unique_resources(mpp_path: str) -> list[str]:
    project = tsk.Project(mpp_path)
    unique = OrderedDict()
    for r in project.resources:
        raw = get_resource_name(r) or ""
        norm = normalize_resource_name(raw)
        if not norm:
            continue
        if norm not in unique:
            unique[norm] = re.sub(r"\s+", " ", raw.replace("\u00A0", " ").strip())
    return list(unique.values())


def parse_indices(input_str: str, max_index: int) -> list[int]:
    tokens = re.findall(r"\d+", input_str)
    seen = set()
    result = []
    for t in tokens:
        try:
            n = int(t)
            if 1 <= n <= max_index and n not in seen:
                seen.add(n)
                result.append(n)
        except Exception:
            continue
    return result


def _to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    s = s.replace("\u00A0", " ")
    s = s.replace(" ", "")
    s = s.replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group(0))
    except Exception:
        return None


def load_bonuses_from_excel(xlsx_path: str) -> tuple[float | None, float | None]:
    if openpyxl is None:
        raise RuntimeError("Для чтения Excel требуется пакет openpyxl. Установите: pip install openpyxl")
    target_col_label = normalize_resource_name("общий итог")
    row_staff_label = normalize_resource_name("премия исполнителей")
    row_manager_label = normalize_resource_name("премия руководителя")
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    staff_bonus = None
    manager_bonus = None
    for ws in wb.worksheets:
        positions = {}
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.strip():
                    key = normalize_resource_name(val)
                    if key and key not in positions:
                        positions[key] = cell.coordinate
        if target_col_label in positions:
            col = openpyxl.utils.cell.column_index_from_string(
                re.sub(r"\d+", "", positions[target_col_label])
            )
            if row_staff_label in positions and staff_bonus is None:
                row_idx = int(re.sub(r"\D+", "", positions[row_staff_label]))
                v = ws.cell(row=row_idx, column=col).value
                staff_bonus = _to_number(v)
            if row_manager_label in positions and manager_bonus is None:
                row_idx = int(re.sub(r"\D+", "", positions[row_manager_label]))
                v = ws.cell(row=row_idx, column=col).value
                manager_bonus = _to_number(v)
        if staff_bonus is not None and manager_bonus is not None:
            break
    return staff_bonus, manager_bonus


def main():
    if len(sys.argv) != 4:
        print("Использование: python -m bonus_calculator <путь_к_файлу.mpp> <путь_к_файлу.xlsx> <путь_для_отчета>")
        sys.exit(2)
    mpp_arg = sys.argv[1]
    xlsx_arg = sys.argv[2]
    report_path = sys.argv[3]
    def resolve_file(p: str) -> str:
        p1 = os.path.abspath(p)
        if os.path.isfile(p1):
            return p1
        base = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        p2 = os.path.abspath(os.path.join(base, p))
        if os.path.isfile(p2):
            return p2
        return p1
    mpp_path = resolve_file(mpp_arg)
    xlsx_path = resolve_file(xlsx_arg)
    if not os.path.isfile(mpp_path) or not mpp_path.lower().endswith(".mpp"):
        print("Ошибка: первый аргумент должен быть существующим файлом с расширением .mpp")
        sys.exit(2)
    if not os.path.isfile(xlsx_path) or not xlsx_path.lower().endswith(".xlsx"):
        print("Ошибка: второй аргумент должен быть существующим файлом с расширением .xlsx")
        sys.exit(2)
    if not os.path.exists(report_path):
        try:
            os.makedirs(report_path, exist_ok=True)
        except Exception:
            pass
    try:
        staff_bonus, manager_bonus = load_bonuses_from_excel(xlsx_path)
        staff_bonus, manager_bonus = round(staff_bonus, 2), round(manager_bonus, 2)
        print("Премии из Excel:")
        print(f"Премия исполнителей (общий итог): {staff_bonus}")
        print(f"Премия руководителя (общий итог): {manager_bonus}")
    except Exception as e:
        print(f"Ошибка чтения премий из Excel: {e}")
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
    selection = input("Введите номера необходимых ресурсов через запятую: ")
    indices = parse_indices(selection, len(resources))
    chosen = [resources[i - 1] for i in indices]
    if chosen:
        print("Выбранные ресурсы:")
        for c in chosen:
            print(c)
    else:
        print("Не выбрано ни одного ресурса.")


if __name__ == "__main__":
    main()

