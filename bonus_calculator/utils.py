import re

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
