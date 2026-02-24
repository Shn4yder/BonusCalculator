import sys
import re
import datetime
from collections import OrderedDict, defaultdict
try:
    import aspose.tasks as tsk
except Exception as e:
    print("Не удалось импортировать библиотеку aspose-tasks. Установите пакет: pip install aspose-tasks")
    sys.exit(1)

from bonus_calculator.utils import normalize_resource_name

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

def parse_duration_to_hours(val_str):
    # Простейший парсер для строк вида PT8H0M0S
    if not val_str:
        return 0.0
    
    val_str = str(val_str).strip()
    
    try:
        import re
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
            return float(clean_str)

        # Fallback: try float directly (handle comma)
        return float(val_str.replace(",", "."))
    except Exception:
        return 0.0

def get_project_name(project):
    project_name = ""
    try:
        name_0 = ""
        name_1 = ""

        # Get Root Task (ID 0) Name
        if project.root_task and project.root_task.name:
            name_0 = project.root_task.name
            
        # Get Task ID 1 Name
        if project.root_task:
            for child in project.root_task.children:
                if child.id == 1:
                    if child.name:
                        name_1 = child.name
                    break
        
        # Select the longest name
        if len(name_1) > len(name_0):
            project_name = name_1
        else:
            project_name = name_0

    except Exception:
        pass
        
    if not project_name:
         # Fallback to Subject/Title
         try:
             if hasattr(project, "subject") and project.subject:
                 project_name = project.subject
             elif hasattr(project, "title") and project.title:
                 project_name = project.title
         except:
             pass
             
    if not project_name:
         project_name = ""
         
    return project_name

def is_project_completed(mpp_path: str) -> bool:
    project = tsk.Project(mpp_path)
    
    # We rely on MS Project's own calculation of the project completion.
    # The Root Task (Project Summary Task) percentage reflects the overall status,
    # automatically handling task durations, rollups, and ignoring inactive tasks.
    if project.root_task:
        # Check if project is 100% complete
        return project.root_task.percent_complete == 100
        
    return False

def collect_timephased_data(mpp_path, chosen_resources):
    # Инициализируем проект заново для работы с объектами
    project = tsk.Project(mpp_path)
    
    # Сопоставим выбранные имена с объектами ресурсов
    norm_map = {}
    for r in project.resources:
        raw = get_resource_name(r)
        norm = normalize_resource_name(raw)
        if norm and norm not in norm_map:
            norm_map[norm] = r
            
    selected_resources = []
    for name in chosen_resources:
        norm = normalize_resource_name(name)
        if norm in norm_map:
            selected_resources.append((name, norm_map[norm]))

    res_data = {}
    
    # Определяем диапазон дат проекта для заголовков месяцев
    project_start = project.start_date
    project_finish = project.finish_date
    
    sorted_months = []
    if project_start and project_finish:
        # Приводим к началу месяца для корректного сравнения
        curr = datetime.datetime(project_start.year, project_start.month, 1)
        end_dt = datetime.datetime(project_finish.year, project_finish.month, 1)
        while curr <= end_dt:
            sorted_months.append((curr.year, curr.month))
            if curr.month == 12:
                curr = datetime.datetime(curr.year + 1, 1, 1)
            else:
                curr = datetime.datetime(curr.year, curr.month + 1, 1)
    
    # Map resource UID to our selected resource index
    uid_to_index = {}
    for idx, (name, r) in enumerate(selected_resources):
        try:
            uid = None
            if hasattr(r, "uid"):
                uid = r.uid
            elif hasattr(r, "get"):
                uid = r.get(tsk.Rsc.UID)
            if uid is not None:
                uid_to_index[uid] = idx
        except Exception:
            pass

    # Initialize res_data
    for idx in range(len(selected_resources)):
        res_data[idx] = defaultdict(float)

    # Iterate assignments to aggregate work
    for ra in project.resource_assignments:
        try:
            r = ra.resource
            if r is None:
                continue
            
            r_uid = None
            if hasattr(r, "uid"):
                r_uid = r.uid
            elif hasattr(r, "get"):
                r_uid = r.get(tsk.Rsc.UID)
                
            if r_uid is not None and r_uid in uid_to_index:
                idx = uid_to_index[r_uid]
                
                # Use a wide date range to ensure we capture all actual work, 
                # even if it falls outside the project's configured start/finish dates.
                # Some tasks might be scheduled manually or have actuals outside the project summary range.
                start_check = datetime.datetime(2000, 1, 1)
                finish_check = datetime.datetime(2050, 1, 1)
                
                td_collection = ra.get_timephased_data(start_check, finish_check, tsk.TimephasedDataType.ASSIGNMENT_ACTUAL_WORK)
                
                for td in td_collection:
                    val = td.value
                    hours = parse_duration_to_hours(val)
                    if hours > 0:
                        dt = td.start
                        key = (dt.year, dt.month)
                        res_data[idx][key] += hours
        except Exception as e:
            pass

    # Recalculate sorted_months based on actual data found to ensure everything is covered
    all_yms = set()
    
    # Add project range months if available (to ensure minimum coverage)
    if project_start and project_finish:
        # Use timezone-naive for comparison logic below
        ps = project_start.replace(tzinfo=None)
        pf = project_finish.replace(tzinfo=None)
        
        curr = datetime.datetime(ps.year, ps.month, 1)
        end_dt = datetime.datetime(pf.year, pf.month, 1)
        while curr <= end_dt:
            all_yms.add((curr.year, curr.month))
            if curr.month == 12:
                curr = datetime.datetime(curr.year + 1, 1, 1)
            else:
                curr = datetime.datetime(curr.year, curr.month + 1, 1)
                
    # Add actual data months
    for idx in res_data:
        for ym in res_data[idx]:
            all_yms.add(ym)
            
    # Sort months
    sorted_months = sorted(list(all_yms))
    
    # Fill gaps? (e.g. if we have Jan and Mar, we should probably have Feb)
    if sorted_months:
        min_ym = sorted_months[0]
        max_ym = sorted_months[-1]
        
        start_dt = datetime.datetime(min_ym[0], min_ym[1], 1)
        end_dt = datetime.datetime(max_ym[0], max_ym[1], 1)
        
        filled_months = []
        curr = start_dt
        while curr <= end_dt:
            filled_months.append((curr.year, curr.month))
            if curr.month == 12:
                curr = datetime.datetime(curr.year + 1, 1, 1)
            else:
                curr = datetime.datetime(curr.year, curr.month + 1, 1)
        sorted_months = filled_months
            
    return project, selected_resources, sorted_months, res_data
