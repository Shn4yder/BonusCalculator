import sys
import os
import re

try:
    import aspose.tasks as tsk
except ImportError:
    print("Please install aspose-tasks: pip install aspose-tasks")
    sys.exit(1)

def parse_duration_string(val_str):
    """
    Парсит строку длительности, которую может вернуть aspose.tasks.
    Поддерживает форматы:
    - PT8H0M0S (ISO 8601 duration)
    - 1209,9 hrs (строковое представление с единицами)
    """
    if not val_str:
        return 0.0
    
    val_str = str(val_str).strip()
    
    # Format: PT...
    m = re.search(r'PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(\.\d+)?)S)?', val_str)
    if m:
        h = int(m.group(1) or 0)
        mm = int(m.group(2) or 0)
        s = float(m.group(3) or 0)
        return h + mm / 60.0 + s / 3600.0
    
    # Format: "1209,9 hrs" or similar
    if "hrs" in val_str.lower():
        clean_str = val_str.lower().replace("hrs", "").replace(",", ".").strip()
        try:
            return float(clean_str)
        except ValueError:
            pass
            
    # Fallback: try float directly
    try:
        return float(val_str.replace(",", "."))
    except ValueError:
        return 0.0

def main():
    file_path = os.path.abspath(r"data\input\input_project.mpp")
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return

    try:
        print(f"Loading project from: {file_path}")
        project = tsk.Project(file_path)
        
        # Получаем корневую задачу (Project Summary Task)
        root_task = project.root_task
        
        if root_task is None:
            print("Root task not found.")
            return

        # Получаем Actual Work (Фактические трудозатраты)
        hours = 0.0
        actual_work_obj = None
        
        try:
            actual_work_obj = root_task.actual_work
            
            # Попытка получить значение через TimeSpan
            if hasattr(actual_work_obj, "time_span"):
                ts = actual_work_obj.time_span
                # В pythonnet свойства обычно доступны как атрибуты
                if hasattr(ts, "total_hours"):
                     hours = ts.total_hours
                elif hasattr(ts, "TotalHours"):
                     hours = ts.TotalHours
            else:
                # Если time_span недоступен, пробуем парсить строку
                hours = parse_duration_string(str(actual_work_obj))
                
            # Если hours все еще 0, пробуем to_double или повторный парсинг строки
            if hours == 0.0:
                if hasattr(actual_work_obj, "to_double"):
                    val = actual_work_obj.to_double()
                    if val > 0:
                        # Предполагаем, что значение в часах, если оно совпадает с ожидаемым порядком
                        hours = val
                
                # Финальная попытка - парсинг строки, если предыдущие методы не сработали
                if hours == 0.0:
                    hours = parse_duration_string(str(actual_work_obj))

        except AttributeError as e:
             # Если доступ к свойствам не удался
             if actual_work_obj:
                 hours = parse_duration_string(str(actual_work_obj))
             else:
                 print(f"Could not access actual_work property: {e}")
        
        print(f"Сумма фактических трудозатрат: {hours} часов")

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
