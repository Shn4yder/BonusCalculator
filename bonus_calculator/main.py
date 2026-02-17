import sys
import os
import re
from collections import OrderedDict
try:
    import aspose.tasks as tsk
except Exception as e:
    print("Не удалось импортировать библиотеку aspose-tasks. Установите пакет: pip install aspose-tasks")
    sys.exit(1)


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


def main():
    if len(sys.argv) != 4:
        print("Использование: python -m bonus_calculator <путь_к_файлу.mpp> <путь_к_файлу.xlsx> <путь_для_отчета>")
        sys.exit(2)
    mpp_path = sys.argv[1]
    xlsx_path = sys.argv[2]
    report_path = sys.argv[3]
    if not os.path.isfile(mpp_path) or not mpp_path.lower().endswith(".mpp"):
        print("Ошибка: первый аргумент должен быть существующим файлом с расширением .mpp")
        sys.exit(2)
    if not os.path.isfile(xlsx_path) or not xlsx_path.lower().endswith(".xlsx"):
        print("Ошибка: второй аргумент должен быть существующим файлом с расширением .xlsx")
        sys.exit(2)
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

