import pdfplumber
import json
from config import DAYS, SLOT_X_AXIS, TIME_LABELS
from helpers import get_duration_hours, extract_instructor, clean_and_split_course, _time_range

import pandas as pd
import xlsxwriter
import re 
from collections import Counter

class PDFScheduleProcessor:
    
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.all_schedules = []

    def _detect_slots(self, day_lines, day, schedule, page):
        xs = sorted({round(line["x0"], 2) for line in day_lines})
        
        # Start of Bug Fix
        """
        When the x0 lines are extracted from the given page, 
        it start from 117.72 to 779.76, but we have the starting from the 88.94.
        As, this is not available in the xs lines list so the 8:30 - 9:00 lecture 
        is registered as booked for each day even if there is not. To fix it after 
        Detecting x0 lines add 88.94 in the first index.
        """
        xs.insert(1, 88.94)
        # End of bug Fix
        
        lecture_no = 1
        start_idx = None

        #print("XS :: ", xs)
        for idx, x in enumerate(SLOT_X_AXIS):
            #print("X = ", x, "  Idx = ", idx)
            if x not in xs:
                if start_idx is None:
                    start_idx = idx - 1 if idx - 1 >= 0 else idx
            else:
                if start_idx is not None:
                    #print("Start Index: ", start_idx)
                    self._add_lecture(start_idx, idx, day, schedule, page, lecture_no)
                    lecture_no += 1
                    start_idx = None

        if start_idx is not None:
            self._add_lecture(start_idx, len(SLOT_X_AXIS) - 1, day, schedule, page, lecture_no)

    def _add_lecture(self, start_idx, end_idx, day, schedule, page, lecture_no):
        start_time = TIME_LABELS[start_idx]
        end_time = TIME_LABELS[end_idx]
        dur = get_duration_hours(start_time, end_time)
        y0, y1 = DAYS[day]
        x0, x1 = SLOT_X_AXIS[start_idx], SLOT_X_AXIS[end_idx]

        cropped = page.crop((x0, y0, x1, y1))
        
        words = cropped.extract_words(keep_blank_chars=False)

        if not words:
            course_info = {"CourseName": "Unknown", "Class": "Unknown", "Venue": "Unknown"}
        
        else:
            # --- 1. Manually rebuild lines from words ---
            lines = []
            current_line = []
            current_top = -1
            
            sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
            
            for w in sorted_words:
                if current_top == -1: 
                    current_top = w['top']
                
                if abs(w['top'] - current_top) > 2:
                    if current_line: 
                        lines.append(current_line)
                    current_line = [w] 
                    current_top = w['top']
                else:
                    current_line.append(w)
            
            if current_line: 
                lines.append(current_line)

            # --- 2. Apply robust parsing logic ---
            
            if not lines:
                 course_info = {"CourseName": "Unknown", "Class": "Unknown", "Venue": "Unknown"}
            
            else:
                line_texts = [" ".join([w['text'] for w in line]) for line in lines]

                # --- This regex is our anchor for finding classes ---
                class_id_pattern = re.compile(r"^(FA|SP)\d{2}-")

                class_start_index = -1
                for i, line in enumerate(line_texts):
                    # Use .search() to find the pattern anywhere in the string
                    # This handles cases where text is fused, like "MAD(G2) FA23..."
                    if class_id_pattern.search(line):
                        class_start_index = i
                        break
                
                course_lines = []
                class_lines_raw = []
                
                if class_start_index == -1:
                    # No class IDs found, assume all lines are the course name
                    course_lines = line_texts
                else:
                    # We found a class!
                    course_lines = line_texts[0:class_start_index]
                    class_lines_raw = line_texts[class_start_index:]

                # --- 3. Separate Venue from end of Class List ---
                venue_text = "Unknown"
                class_lines_final = []

                if class_lines_raw:
                    last_line = class_lines_raw[-1]
                    
                    # If the last line does NOT look like a class, it's a venue
                    if not class_id_pattern.search(last_line):
                        venue_text = last_line
                        class_lines_final = class_lines_raw[:-1] # All but the last
                    else:
                        # Last line is a class, no venue
                        venue_text = "Unknown"
                        class_lines_final = class_lines_raw
                
                # --- 4. Post-Process: Fix Fused Lines (Your Edge Case) ---
                # This fixes "SP23-BCS-B/N-4"
                
                # This pattern matches: (Group 1: Class ID) (Group 2: Venue)
                venue_fuse_pattern = re.compile(r"^((?:FA|SP)\d{2}-[\w/-]+/?)\s+([\w\-() ]+)$")
                
                processed_class_lines = []
                for line in class_lines_final:
                    match = venue_fuse_pattern.match(line)
                    if match:
                        # It's a fused line!
                        processed_class_lines.append(match.group(1).strip()) # Add "SP23-BCS-B/"
                        venue_text = match.group(2).strip() # Set "N-4" as venue
                    else:
                        # It's a normal line
                        processed_class_lines.append(line)

                # --- 5. Assemble final dictionary ---
                course_info = {
                    "CourseName": " ".join(course_lines) if course_lines else "Unknown",
                    "Class": "\n".join(processed_class_lines),
                    "Venue": venue_text
                }

        # --- This part remains the same ---
        lecture = {
            "Lecture": lecture_no,
            "start-time-slot": SLOT_X_AXIS[start_idx],
            "end-time-slot": SLOT_X_AXIS[end_idx],
            "start-time": start_time,
            "end-time": end_time,
            "duration": str(dur),
            "type": "Lab" if dur == 3.0 else "Lecture",
            **course_info  # Add our newly structured data
        }
        schedule[day].append(lecture)
              
    def _get_vertical_lines(self, page):
        day_lines = {day: [] for day in DAYS}
        for obj in page.objects.get("line", []):
            if abs(obj["x0"] - obj["x1"]) < 0.5:  # vertical line
                (x0, y0), (x1, y1) = obj["pts"]
                for day, bounds in DAYS.items():
                    inside = min(y0, y1) >= min(bounds) - 1 and max(y0, y1) <= max(bounds) + 1
                    if inside:
                        line_data = {
                            "x0": round(x0, 2),
                            "x1": round(x1, 2),
                            "y0": round(y0, 2),
                            "y1": round(y1, 2)
                        }
                        day_lines[day].append(line_data)
        return day_lines

    def process_pdf(self):
        with pdfplumber.open(self.pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                
                print("Page # ", page_num)
                
                instructor_text = extract_instructor(page.extract_text())

                # Skip non-CS teachers
                if not instructor_text.startswith("CS-"):
                    continue
                
                print("Page # ", page_num)
                
                schedule = {
                    "Page": page_num,
                    "Instructor": extract_instructor(page.extract_text()),
                    "Monday": [], "Tuesday": [], "Wednesday": [], "Thursday": [], "Friday": []
                }

                day_lines = self._get_vertical_lines(page)
                for day in DAYS:
                    self._detect_slots(day_lines[day], day, schedule, page)

                self.all_schedules.append(schedule)
        return self.all_schedules

    def export_to_json(self, output_path):
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(self.all_schedules, f, ensure_ascii=False, indent=4)    
            
    def export_to_excel(self, output_excel):
        # Use an ordered list to maintain sheet order
        day_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        
        # 1. Create a dictionary to hold rows for each day
        data_by_day = {day: [] for day in day_list}

        for record in self.all_schedules:
            teacher = record["Instructor"]
            dept = teacher.split("-")[0] if "-" in teacher else ""
            # Handle cases where there might not be a '-'
            teacher_name = teacher.split("-")[1] if "-" in teacher and len(teacher.split("-")) > 1 else teacher
            print("Processing ", teacher_name)

            for day in day_list:
                for lec in record.get(day, []):
                    course = lec.get("CourseName", "")
                    cls = course.split()[-2] if len(course.split()) > 1 else ""
                    duration = float(lec.get("duration", 0))
                    room = lec.get("Venue", "Unknown")
                    
                    start_time_str = lec["start-time"]
                    end_time_str = lec["end-time"]

                    # Base row data, common to all rows for this lecture
                    base_row_data = {
                        "Teacher Dept": dept,
                        "Teacher": teacher_name,
                        "Subject": course,
                        "Classes": cls,
                        "Day": day, 
                        "Room": room
                    }

                    if duration == 3.0:
                        # --- Split 3.0hr lab into two 1.5hr slots ---
                        
                        # Calculate midpoint time (1.5 hours after start)
                        start_h, start_m = map(int, start_time_str.split(':'))
                        mid_m_total = start_m + 90  # 90 minutes = 1.5 hours
                        mid_h = start_h + (mid_m_total // 60)
                        mid_m = mid_m_total % 60
                        mid_time_str = f"{mid_h:02d}:{mid_m:02d}"

                        # Row 1 (First 1.5hr)
                        row_1 = base_row_data.copy()
                        row_1.update({
                            "Period Length": 1.5,
                            "Period": 1,
                            "Period_Time": f"{start_time_str}-{mid_time_str}",
                        })
                        data_by_day[day].append(row_1)

                        # Row 2 (Second 1.5hr)
                        row_2 = base_row_data.copy()
                        row_2.update({
                            "Period Length": 1.5,
                            "Period": 2,
                            "Period_Time": f"{mid_time_str}-{end_time_str}",
                        })
                        data_by_day[day].append(row_2)
                        
                    else:
                        # --- Handle other lectures (1.5hr, 1.0hr, etc.) as one block ---
                        row_1 = base_row_data.copy()
                        row_1.update({
                            "Period Length": duration,
                            "Period": 1,
                            "Period_Time": f"{start_time_str}-{end_time_str}",
                        })
                        data_by_day[day].append(row_1)

        # --- This is the new, fully custom styling section ---

        # 2. Create the writer and workbook (once)
        writer = pd.ExcelWriter(output_excel, engine='xlsxwriter')
        workbook = writer.book

        # 3. Define your custom formats (once, can be reused)
        
        # --- Header Format ---
        header_format = workbook.add_format({
            'bold': True,
            'font_color': '#FFFFFF',
            'bg_color': '#000000', # Black
            'align': 'center',
            'valign': 'vcenter',
            'bottom': 2, # Medium-thick border
            'border_color': '#000000'
        })

        # --- Body Format (Left Aligned) ---
        body_left_format = workbook.add_format({
            'bg_color': '#FFFFFF', # White
            'border': 1,
            'border_color': '#BFBFBF', # Light gray
            'align': 'left',
            'valign': 'vcenter'
        })

        # --- Body Format (Center Aligned) ---
        body_center_format = workbook.add_format({
            'bg_color': '#FFFFFF',
            'border': 1,
            'border_color': '#BFBFBF',
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # 4. Define column order and centered columns
        column_order = [
            "Teacher Dept", "Teacher", "Subject", "Classes",
            "Period Length", "Period", "Period_Time", "Room"
        ]
        centered_cols = {'Period Length', 'Period', "Period_Time", "Teacher Dept"}

        # 5. Loop through each day and write to a separate sheet
        for day in day_list:
            rows_for_this_day = data_by_day[day]
            
            if not rows_for_this_day:
                print(f"No data for {day}, skipping sheet.")
                continue
            
            print(f"Writing sheet for {day}...")

            # Add a new worksheet for the day
            worksheet = workbook.add_worksheet(day)

            # --- Apply all styling to this new sheet ---
            
            # A. Set Column Widths
            worksheet.set_column('A:A', 13) # Teacher Dept
            worksheet.set_column('B:B', 30) # Teacher
            worksheet.set_column('C:C', 60) # Subject
            worksheet.set_column('D:D', 25) # Classes
            worksheet.set_column('E:E', 12) # Period Length
            # worksheet.set_column('F:F', 15) # Day
            worksheet.set_column('F:F', 10) # Period
            worksheet.set_column('G:G', 17) # Period_Time
            worksheet.set_column('H:H', 45) # Room

            # B. Write the Header Row
            worksheet.set_row(0, 30) # Set header row height
            worksheet.write_row('A1', column_order, header_format)

            # C. Write the Data Rows
            for row_num, row_data in enumerate(rows_for_this_day, start=1):
                for col_num, col_name in enumerate(column_order):
                    value = row_data.get(col_name)
                    
                    # Pick the correct format
                    if col_name in centered_cols:
                        current_format = body_center_format
                    else:
                        current_format = body_left_format
                    
                    # Write the individual cell
                    worksheet.write(row_num, col_num, value, current_format)

        # 6. Close and save the file (outside the loop)
        writer.close()
        
        print(f"Successfully exported custom-styled Excel to {output_excel}")
    
    