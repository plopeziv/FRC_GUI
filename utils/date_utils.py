from datetime import datetime

def parse_string_date(date):
    if not date:
        return None
    
    date = str(date).strip()
    
    formats = (
        "%m/%d/%y",
        "%m/%d/%Y",
        "%m-%d-%y",
        "%m-%d-%Y",
    )

    for fmt in formats:
        try:
            return datetime.strptime(date, fmt)
        except ValueError:
            pass

    return None