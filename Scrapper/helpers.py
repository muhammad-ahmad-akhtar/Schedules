import re
from datetime import datetime, timedelta

def get_duration_hours(start_str, end_str):
    fmt = "%H:%M"
    start = datetime.strptime(start_str, fmt)
    end = datetime.strptime(end_str, fmt)
    return (end - start).total_seconds() / 3600

def extract_instructor(text):
    if not text:
        return "Not Mentioned"
    for line in text.splitlines():
        if line.strip().startswith("Teacher"):
            match = re.search(r"Teacher\s+(.*)", line)
            if match:
                return match.group(1).strip()
    return "Not Mentioned"

def clean_and_split_course(text):
    if not text:
        return {"CourseName": "Unknown", "Class": "Unknown", "Venue": "Unknown"}
    
    print(text)

    clean_text = re.sub(r"\s+", " ", text).strip()
    pattern = r"^(.*?)\s+(FA\d{2}-[A-Z]+-\w)\s+(.+)$"
    match = re.match(pattern, clean_text)

    if match:
        return {
            "CourseName": match.group(1).strip(),
            "Class": match.group(2).strip(),
            "Venue": match.group(3).strip()
        }
    return {"CourseName": clean_text, "Class": "Unknown", "Venue": "Unknown"}

def _time_range(start, end, step_minutes=30):
        start_t = datetime.strptime(start, "%H:%M")
        end_t = datetime.strptime(end, "%H:%M")
        while start_t < end_t:
            next_t = start_t + timedelta(minutes=step_minutes)
            yield f"{start_t.strftime('%H:%M')} - {next_t.strftime('%H:%M')}"
            start_t = next_t

