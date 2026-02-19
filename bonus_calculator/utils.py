import re
import os
import datetime

RUSSIAN_MONTHS = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май", 6: "Июнь",
    7: "Июль", 8: "Август", 9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

def normalize_resource_name(name: str) -> str:
    if name is None:
        return ""
    s = str(name).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s.strip())
    return s.casefold()

def to_number(value):
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

def sanitize_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[<>:\"/\\\\|?*]', '', name)
    name = re.sub(r'\\s+', ' ', name)
    return name

def get_unique_report_path(base_dir_or_file: str, project_name: str = None) -> str:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if project_name:
        safe_proj_name = sanitize_filename(project_name)
        if len(safe_proj_name) > 50:
            safe_proj_name = safe_proj_name[:50]
        base_filename = f"Отчет_{safe_proj_name}_{timestamp}.xlsx"
    else:
        base_filename = f"Отчет_{timestamp}.xlsx"
    
    target_dir = "."
    target_filename = base_filename
    
    if os.path.isdir(base_dir_or_file):
        target_dir = base_dir_or_file
        target_filename = base_filename
    elif base_dir_or_file.lower().endswith('.xlsx'):
        dirname = os.path.dirname(base_dir_or_file)
        basename = os.path.basename(base_dir_or_file)
        if dirname:
            target_dir = dirname
        else:
            target_dir = "."
            
        if os.path.exists(base_dir_or_file):
            # File exists, append timestamp
            name, ext = os.path.splitext(basename)
            target_filename = f"{name}_{timestamp}{ext}"
        else:
            # File does not exist, use as is (but ensure dir exists)
            target_filename = basename
    else:
        # Fallback
        target_dir = base_dir_or_file
        target_filename = base_filename
        
    if not os.path.exists(target_dir):
        try:
            os.makedirs(target_dir, exist_ok=True)
        except:
            pass
            
    return os.path.join(target_dir, target_filename)
