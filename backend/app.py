from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.utils import get_column_letter
import io
from datetime import datetime
import unicodedata
from copy import copy
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas
import re
import csv

app = Flask(__name__)
CORS(app)

# ── helpers ────────────────────────────────────────────────────────────────
def normalize_str(s):
    if s is None:
        return ""
    s_str = str(s).strip()
    if s_str.startswith('='):
        return ""
    # Remove accents and lowercase
    s_str = s_str.lower()
    s_str = "".join(
        c for c in unicodedata.normalize('NFD', s_str)
        if unicodedata.category(c) != 'Mn'
    )
    return s_str

def is_csv_bytes(file_bytes):
    if file_bytes.startswith(b'PK\x03\x04'):
        return False
    if file_bytes.startswith(b'\xd0\xcf\x11\xe0'):
        return False
    return True

def parse_csv_to_rows(file_bytes):
    text = None
    for encoding in ['utf-8-sig', 'utf-8', 'iso-8859-1']:
        try:
            text = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("Impossible de décoder le fichier CSV.")
    
    first_line = text.split('\n')[0] if text else ""
    delimiter = ','
    if ';' in first_line:
        delimiter = ';'
    elif '\t' in first_line:
        delimiter = '\t'
        
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    return list(reader)

def load_workbook_from_bytes(file_bytes, data_only=False):
    if is_csv_bytes(file_bytes):
        rows = parse_csv_to_rows(file_bytes)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CSV Data"
        for r_idx, row in enumerate(rows, 1):
            for c_idx, val in enumerate(row, 1):
                if val is not None:
                    val_str = str(val).strip()
                    if val_str.isdigit():
                        ws.cell(row=r_idx, column=c_idx, value=int(val_str))
                    else:
                        try:
                            ws.cell(row=r_idx, column=c_idx, value=float(val_str.replace(',', '.')))
                        except ValueError:
                            ws.cell(row=r_idx, column=c_idx, value=val_str)
                else:
                    ws.cell(row=r_idx, column=c_idx, value="")
        return wb
    else:
        return openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=data_only)

def format_single_name(name_str):
    if not name_str:
        return ""
    words = name_str.split()
    if not words:
        return ""
    if len(words) == 1:
        return words[0].title()
    
    first_word = words[0]
    first_is_upper = first_word.isupper() and len(first_word) > 1
    others_all_upper = all(w.isupper() for w in words[1:])
    
    if first_is_upper and not others_all_upper:
        idx = 0
        while idx < len(words) and words[idx].isupper():
            idx += 1
        if idx < len(words):
            prenom_part = " ".join(words[idx:])
            nom_part = " ".join(words[:idx])
            return f"{prenom_part.title()} {nom_part.upper()}"
            
    idx = len(words) - 1
    while idx >= 0 and words[idx].isupper():
        idx -= 1
    
    if idx >= 0 and idx < len(words) - 1:
        prenom_part = " ".join(words[:idx+1])
        nom_part = " ".join(words[idx+1:])
        return f"{prenom_part.title()} {nom_part.upper()}"
        
    return f"{words[0].title()} {' '.join(words[1:]).upper()}"

def format_employee_name(prenom_val, nom_val):
    if prenom_val is not None and nom_val is not None:
        p_str = str(prenom_val).strip()
        n_str = str(nom_val).strip()
        if p_str and n_str:
            return f"{p_str.title()} {n_str.upper()}"
        elif p_str:
            return format_single_name(p_str)
        elif n_str:
            return format_single_name(n_str)
    elif nom_val is not None:
        return format_single_name(str(nom_val).strip())
    elif prenom_val is not None:
        return format_single_name(str(prenom_val).strip())
    return ""

def copy_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = copy(src_cell.number_format)

def detect_structure(ws):
    EMPLOYEE_KEYWORDS = {
        "employe", "employee", "employes", "employees",
        "nom", "name", "agent", "collaborateur", "personnel",
        "user", "salarie", "travailleur", "staff",
        "prenom", "nom et prenom", "nom & prenom", "nom prenom",
        "collaborateurs", "agents", "salaries"
    }

    SUMMARY_KEYWORDS = {
        "total", "totaux", "somme", "sum", "moyenne", "average", "global"
    }

    EXCLUDED_HEADERS = {
        "departement", "service", "role", "poste", "fonction", "adresse",
        "telephone", "tel", "email", "mail", "status", "embauche", "contrat",
        "commentaire", "comment", "note", "observation", "remarque", "sexe",
        "genre", "age", "date", "jour", "month", "mois", "annee", "year",
        "temps", "time", "timestamp", "horodatage"
    }

    max_r = min(30, ws.max_row)
    max_c = min(30, ws.max_column)

    candidates = []
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            val = ws.cell(row=r, column=c).value
            if val is not None:
                val_str = str(val).strip()
                if val_str.startswith('='):
                    continue
                val_norm = normalize_str(val_str)
                
                # Exclude header candidates that match total/summary words
                if any(k in val_norm for k in ["total", "somme", "sum", "totaux", "average", "moyenne"]):
                    continue
                    
                is_employee_cell = False
                if val_norm in EMPLOYEE_KEYWORDS:
                    is_employee_cell = True
                elif val_norm in {"nom", "name", "id", "employe", "employee"}:
                    is_employee_cell = True
                elif any(k in val_norm for k in ["employe", "employee", "collaborateur", "agent", "salarie", "staff", "personnel"]):
                    is_employee_cell = True
                    
                if is_employee_cell:
                    priority = 2 if any(k in val_norm for k in ["employe", "employee", "nom", "name", "collaborateur"]) else 1
                    candidates.append((r, c, priority))

    if candidates:
        candidates.sort(key=lambda x: (-x[2], x[0], x[1]))
        header_row, emp_col, _ = candidates[0]
    else:
        header_row = None
        emp_col = None
        for c in range(1, max_c + 1):
            text_count = 0
            num_count = 0
            first_text_row = None
            for r in range(1, min(15, ws.max_row) + 1):
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    if isinstance(val, (int, float)):
                        num_count += 1
                    elif isinstance(val, str) and val.strip():
                        if val.strip().startswith('='):
                            num_count += 1
                        else:
                            text_count += 1
                            if first_text_row is None:
                                first_text_row = r
            if text_count > 2 and text_count > num_count:
                header_row = first_text_row
                emp_col = c
                break

        if header_row is None:
            header_row = 1
            emp_col = 1

    # Detect prenom_col (first name column) if separate from emp_col (last name)
    prenom_col = None
    if header_row is not None:
        for c in range(1, max_c + 1):
            val = ws.cell(row=header_row, column=c).value
            if val is not None:
                val_norm = normalize_str(val)
                if "prenom" in val_norm:
                    prenom_col = c
                    break
        
        if prenom_col is not None and emp_col == prenom_col:
            found_other = False
            for c in range(1, max_c + 1):
                if c == prenom_col:
                    continue
                val = ws.cell(row=header_row, column=c).value
                if val is not None:
                    val_norm = normalize_str(val)
                    if any(k in val_norm for k in ["nom", "name", "employe", "employee", "collaborateur", "agent", "salarie"]):
                        emp_col = c
                        found_other = True
                        break
            if not found_other:
                prenom_col = None

    # Find employee rows and summary rows
    emp_rows = []
    total_rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        emp_val = ws.cell(row=r, column=emp_col).value
        if emp_val is None:
            continue
        
        val_str = str(emp_val).strip()
        if not val_str:
            continue
            
        if val_str.startswith('='):
            val_norm = val_str.lower()
            is_summary = any(k in val_norm for k in SUMMARY_KEYWORDS)
            if is_summary:
                total_rows.append(r)
            else:
                emp_rows.append(r)
        else:
            emp_str = normalize_str(emp_val)
            if emp_str in SUMMARY_KEYWORDS:
                total_rows.append(r)
            else:
                emp_rows.append(r)

    # Find columns to sum
    sum_cols = []
    existing_total_col = None

    for c in range(emp_col + 1, ws.max_column + 1):
        if prenom_col is not None and c == prenom_col:
            continue

        header_val = ws.cell(row=header_row, column=c).value
        header_norm = normalize_str(header_val)
        
        if header_norm in {"total", "somme", "sum", "totaux"}:
            existing_total_col = c
            continue
            
        if header_norm in EXCLUDED_HEADERS:
            continue
            
        # Verify if column is mostly numeric
        text_count = 0
        num_count = 0
        for r in emp_rows:
            val = ws.cell(row=r, column=c).value
            if val is not None:
                if isinstance(val, (int, float)):
                    num_count += 1
                elif isinstance(val, str):
                    val_str = val.strip()
                    if val_str.startswith('='):
                        num_count += 1
                    elif val_str:
                        text_count += 1
                        
        if text_count > 0 and text_count >= num_count:
            continue
            
        sum_cols.append(c)

    return header_row, emp_col, prenom_col, emp_rows, total_rows, sum_cols, existing_total_col


def process_workbook_in_place(wb):
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        header_row, emp_col, prenom_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
        
        if not emp_rows or not sum_cols:
            continue
            
        # Determine layout type
        name_counts = {}
        for r in emp_rows:
            val = ws.cell(row=r, column=emp_col).value
            p_val = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
            key = (str(val).strip() if val is not None else "", str(p_val).strip() if p_val is not None else "")
            if key[0] or key[1]:
                name_counts[key] = name_counts.get(key, 0) + 1
                
        is_list_layout = False
        if len(emp_rows) > 0:
            max_occ = max(name_counts.values()) if name_counts else 0
            if max_occ > 1:
                is_list_layout = True
                
        if is_list_layout:
            # Collapse duplicates and sum values in sum_cols
            unique_emps = []
            emp_to_rows = {}
            for r in emp_rows:
                val = ws.cell(row=r, column=emp_col).value
                p_val = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                if val is not None or p_val is not None:
                    nom_str = str(val).strip() if val is not None else ""
                    prenom_str = str(p_val).strip() if p_val is not None else ""
                    key = (nom_str, prenom_str)
                    if key not in emp_to_rows:
                        unique_emps.append(key)
                        emp_to_rows[key] = []
                    emp_to_rows[key].append(r)
            
            rows_to_delete = []
            for key in unique_emps:
                rows = emp_to_rows[key]
                target_row = rows[0]
                
                # Sum daily columns
                for c in sum_cols:
                    col_sum = 0
                    has_numeric = False
                    for r in rows:
                        v = ws.cell(row=r, column=c).value
                        if v is not None:
                            if isinstance(v, (int, float)):
                                col_sum += v
                                has_numeric = True
                            elif isinstance(v, str):
                                val_str = v.strip()
                                if val_str.isdigit():
                                    col_sum += int(val_str)
                                    has_numeric = True
                    if has_numeric:
                        ws.cell(row=target_row, column=c, value=col_sum)
                
                # Collect duplicate rows to delete
                if len(rows) > 1:
                    rows_to_delete.extend(rows[1:])
            
            # Delete duplicates from bottom to top
            for r in sorted(rows_to_delete, reverse=True):
                ws.delete_rows(r)
                
            # Re-detect structure on the collapsed sheet
            header_row, emp_col, prenom_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
            if not emp_rows or not sum_cols:
                continue

        # Find all existing total columns
        total_cols = []
        for c in range(max(emp_col, prenom_col or 0) + 1, ws.max_column + 1):
            header_val = ws.cell(row=header_row, column=c).value
            header_norm = normalize_str(header_val)
            if header_norm in {"total", "somme", "sum", "totaux"}:
                total_cols.append(c)
                
        # Decide total column
        if total_cols:
            total_col = total_cols[0]
        else:
            total_col = ws.max_column + 1
            hdr_cell = ws.cell(row=header_row, column=total_col, value="Total")
            copy_style(ws.cell(row=header_row, column=sum_cols[-1]), hdr_cell)
            
        total_col_letter = get_column_letter(total_col)
        start_col_letter = get_column_letter(sum_cols[0])
        end_col_letter = get_column_letter(sum_cols[-1])
        
        # Populate totals for employees
        for r in emp_rows:
            row_sum = 0
            for c in sum_cols:
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    if isinstance(v, (int, float)):
                        row_sum += v
                    elif isinstance(v, str):
                        val_str = v.strip()
                        if val_str.isdigit():
                            row_sum += int(val_str)
            cell = ws.cell(row=r, column=total_col, value=row_sum)
            copy_style(ws.cell(row=r, column=sum_cols[-1]), cell)
            
        # Populate totals for summary rows
        if emp_rows:
            for r in total_rows:
                col_sum = 0
                for er in emp_rows:
                    v = ws.cell(row=er, column=total_col).value
                    if isinstance(v, (int, float)):
                        col_sum += v
                cell = ws.cell(row=r, column=total_col, value=col_sum)
                copy_style(ws.cell(row=r, column=sum_cols[-1]), cell)
                
    return wb


def lire_excel_dynamique(file_bytes):
    wb = load_workbook_from_bytes(file_bytes, data_only=True)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row, emp_col, prenom_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
        if emp_rows and sum_cols:
            name_counts = {}
            for r in emp_rows:
                val = ws.cell(row=r, column=emp_col).value
                p_val = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                key = (str(val).strip() if val is not None else "", str(p_val).strip() if p_val is not None else "")
                if key[0] or key[1]:
                    name_counts[key] = name_counts.get(key, 0) + 1
                    
            is_list_layout = False
            if len(emp_rows) > 0:
                max_occ = max(name_counts.values()) if name_counts else 0
                if max_occ > 1:
                    is_list_layout = True
                    
            if not is_list_layout:
                dates = []
                for c in sum_cols:
                    hdr_val = ws.cell(row=header_row, column=c).value
                    dates.append(str(hdr_val) if hdr_val is not None else "")
                    
                employees = []
                data = {}
                for r in emp_rows:
                    nom = ws.cell(row=r, column=emp_col).value
                    prenom = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                    if nom is None and prenom is None:
                        continue
                    
                    nom_formatted = format_employee_name(prenom, nom)
                    if not nom_formatted:
                        continue
                        
                    if nom_formatted not in data:
                        employees.append(nom_formatted)
                        data[nom_formatted] = []
                    
                    vals = []
                    for c in sum_cols:
                        v = ws.cell(row=r, column=c).value
                        try:
                            vals.append(int(v) if v is not None else 0)
                        except:
                            vals.append(0)
                    data[nom_formatted] = vals
                return employees, dates, data
            else:
                dates = []
                for c in sum_cols:
                    hdr_val = ws.cell(row=header_row, column=c).value
                    dates.append(str(hdr_val) if hdr_val is not None else "")
                    
                employees_map = {}
                for r in emp_rows:
                    nom = ws.cell(row=r, column=emp_col).value
                    prenom = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                    if nom is None and prenom is None:
                        continue
                        
                    nom_formatted = format_employee_name(prenom, nom)
                    if not nom_formatted:
                        continue
                        
                    if nom_formatted not in employees_map:
                        employees_map[nom_formatted] = [0] * len(sum_cols)
                        
                    for idx, c in enumerate(sum_cols):
                        v = ws.cell(row=r, column=c).value
                        try:
                            employees_map[nom_formatted][idx] += int(v) if v is not None else 0
                        except:
                            pass
                            
                employees = list(employees_map.keys())
                return employees, dates, employees_map
                
    return [], [], {}

MONTH_NAMES = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
    7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
}

def extraire_mois_annee(ws, dates):
    # 1. Try to find date in sheet title
    title = ws.title
    title_clean = normalize_str(title)
    
    # Check for French month names in sheet title
    for m_num, m_name in MONTH_NAMES.items():
        if normalize_str(m_name) in title_clean:
            year_match = re.search(r'\b(20\d{2})\b', title)
            year = year_match.group(1) if year_match else str(datetime.now().year)
            return f"{m_name} {year}"
            
    # Look for MM/YYYY or MM-YYYY in title
    match = re.search(r'\b(0?[1-9]|1[0-2])[-/](20\d{2})\b', title)
    if match:
        m_num = int(match.group(1))
        year = match.group(2)
        return f"{MONTH_NAMES[m_num]} {year}"

    # 2. Try to find from dates headers
    for d in dates:
        if isinstance(d, datetime):
            return f"{MONTH_NAMES[d.month]} {d.year}"
        if isinstance(d, str):
            # Try to match formats like DD/MM/YYYY or YYYY-MM-DD
            match_slash = re.search(r'\b\d{1,2}/(\d{1,2})/(20\d{2})\b', d)
            if match_slash:
                m_num = int(match_slash.group(1))
                year = match_slash.group(2)
                if m_num in MONTH_NAMES:
                    return f"{MONTH_NAMES[m_num]} {year}"
            
            match_dash = re.search(r'\b(20\d{2})-(\d{1,2})-\d{1,2}\b', d)
            if match_dash:
                year = match_dash.group(1)
                m_num = int(match_dash.group(2))
                if m_num in MONTH_NAMES:
                    return f"{MONTH_NAMES[m_num]} {year}"

    # 3. Try to look at the first few rows/cells of the sheet for any cell containing a month or date
    for r in range(1, min(10, ws.max_row + 1)):
        for c in range(1, min(10, ws.max_column + 1)):
            val = ws.cell(row=r, column=c).value
            if isinstance(val, datetime):
                return f"{MONTH_NAMES[val.month]} {val.year}"
            if isinstance(val, str):
                val_norm = normalize_str(val)
                for m_num, m_name in MONTH_NAMES.items():
                    if normalize_str(m_name) in val_norm:
                        year_match = re.search(r'\b(20\d{2})\b', val)
                        year = year_match.group(1) if year_match else str(datetime.now().year)
                        return f"{m_name} {year}"

    # 4. Fallback to current month and year
    now = datetime.now()
    return f"{MONTH_NAMES[now.month]} {now.year}"

def generer_etiquettes_pdf(employees_data):
    # A4 Dimensions: 595.27 x 841.89 points
    page_width, page_height = A4
    
    # Grid configuration (3 columns, 7 rows)
    cols = 3
    rows = 7
    labels_per_page = cols * rows
    
    # Dimensions in points
    label_width = 178
    label_height = 104
    col_gap = 10
    row_gap = 10
    
    # Calculate margins to center the grid
    margin_x = (page_width - (cols * label_width + (cols - 1) * col_gap)) / 2.0
    margin_y = (page_height - (rows * label_height + (rows - 1) * row_gap)) / 2.0
    
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    # Palette colors (LaitTrack aesthetic: teal primary, light teal/slate background)
    primary_color = HexColor("#0c9f9a")      # Teal
    secondary_color = HexColor("#087b77")    # Dark teal
    border_color = HexColor("#d8e3e7")       # Light grey/blue border
    text_dark = HexColor("#172026")          # Dark text
    badge_bg = HexColor("#eef8f7")           # Soft teal badge background
    
    for index, emp in enumerate(employees_data):
        page_idx = index % labels_per_page
        
        # Calculate row and col (0-indexed)
        col_idx = page_idx % cols
        row_idx = page_idx // cols  # 0 to 6
        
        # X and Y positions
        x = margin_x + col_idx * (label_width + col_gap)
        y = page_height - margin_y - (row_idx + 1) * label_height - row_idx * row_gap
        
        # 1. Draw rounded rectangle border
        c.setStrokeColor(border_color)
        c.setLineWidth(1)
        c.setFillColor(HexColor("#ffffff"))
        c.roundRect(x, y, label_width, label_height, 6, fill=True, stroke=True)
        
        # 2. Draw card header accent (top border line or minor logo/title)
        c.setStrokeColor(HexColor("#e6f2f1"))
        c.setLineWidth(0.5)
        c.line(x + 10, y + label_height - 20, x + label_width - 10, y + label_height - 20)
        
        # Draw "[Period]" tiny brand label
        period = emp.get("period", "Étiquette")
        c.setFillColor(secondary_color)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(x + 12, y + label_height - 14, period)
        
        # 3. Draw Employee Name
        name = emp["nom"]
        
        # Set initial font size
        font_size = 13
        c.setFont("Helvetica-Bold", font_size)
        
        # Adjust font size dynamically based on name length
        name_width = c.stringWidth(name, "Helvetica-Bold", font_size)
        max_name_width = label_width - 24
        while name_width > max_name_width and font_size > 8:
            font_size -= 0.5
            name_width = c.stringWidth(name, "Helvetica-Bold", font_size)
            
        c.setFont("Helvetica-Bold", font_size)
        c.setFillColor(text_dark)
        # Center the name horizontally
        name_x = x + (label_width - name_width) / 2.0
        # Vertically, center it
        c.drawString(name_x, y + label_height / 2.0 + 2, name)
        
        # 4. Draw stylized total badge
        badge_w = 120
        badge_h = 22
        badge_x = x + (label_width - badge_w) / 2.0
        badge_y = y + 14
        
        c.setFillColor(badge_bg)
        c.setStrokeColor(HexColor("#bce2e0"))
        c.setLineWidth(0.75)
        c.roundRect(badge_x, badge_y, badge_w, badge_h, 4, fill=True, stroke=True)
        
        # Text inside badge
        total_text = f"{emp['total']} paquets"
        c.setFillColor(primary_color)
        c.setFont("Helvetica-Bold", 10)
        text_w = c.stringWidth(total_text, "Helvetica-Bold", 10)
        text_x = badge_x + (badge_w - text_w) / 2.0
        text_y = badge_y + (badge_h - 10) / 2.0 + 1.5
        c.drawString(text_x, text_y, total_text)
        
        # If we reached the end of the page and there are more elements, start a new page
        if page_idx == labels_per_page - 1 and index < len(employees_data) - 1:
            c.showPage()
            
    c.save()
    buffer.seek(0)
    return buffer.getvalue()

@app.route("/api/labels", methods=["POST"])
def labels():
    """Génère et retourne le PDF des étiquettes."""
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400
    f = request.files["file"]
    try:
        file_bytes = f.read()
        
        # Load workbook
        wb = load_workbook_from_bytes(file_bytes, data_only=True)
        
        labels_data = []
        
        # Process each sheet
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            header_row, emp_col, prenom_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
            if not emp_rows or not sum_cols:
                continue
                
            # Determine if List Layout (Layout B)
            name_counts = {}
            for r in emp_rows:
                val = ws.cell(row=r, column=emp_col).value
                p_val = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                key = (str(val).strip() if val is not None else "", str(p_val).strip() if p_val is not None else "")
                if key[0] or key[1]:
                    name_counts[key] = name_counts.get(key, 0) + 1
                    
            is_list_layout = False
            if len(emp_rows) > 0:
                max_occ = max(name_counts.values()) if name_counts else 0
                if max_occ > 1:
                    is_list_layout = True
            
            dates = []
            for c in sum_cols:
                hdr_val = ws.cell(row=header_row, column=c).value
                dates.append(str(hdr_val) if hdr_val is not None else "")
                
            # Extract period (month/year) for this sheet
            period = extraire_mois_annee(ws, dates)
            
            # Aggregate employees for this sheet
            sheet_employees = []
            sheet_data = {}
            
            if not is_list_layout:
                # Pivot layout (Layout A)
                for r in emp_rows:
                    nom = ws.cell(row=r, column=emp_col).value
                    prenom = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                    if nom is None and prenom is None:
                        continue
                    
                    nom_formatted = format_employee_name(prenom, nom)
                    if not nom_formatted:
                        continue
                        
                    if nom_formatted not in sheet_data:
                        sheet_employees.append(nom_formatted)
                        sheet_data[nom_formatted] = []
                    
                    vals = []
                    for c in sum_cols:
                        v = ws.cell(row=r, column=c).value
                        try:
                            vals.append(int(v) if v is not None else 0)
                        except:
                            vals.append(0)
                    sheet_data[nom_formatted] = vals
            else:
                # List layout (Layout B)
                for r in emp_rows:
                    nom = ws.cell(row=r, column=emp_col).value
                    prenom = ws.cell(row=r, column=prenom_col).value if prenom_col is not None else None
                    if nom is None and prenom is None:
                        continue
                        
                    nom_formatted = format_employee_name(prenom, nom)
                    if not nom_formatted:
                        continue
                        
                    if nom_formatted not in sheet_data:
                        sheet_employees.append(nom_formatted)
                        sheet_data[nom_formatted] = [0] * len(sum_cols)
                        
                    for idx, c in enumerate(sum_cols):
                        v = ws.cell(row=r, column=c).value
                        try:
                            sheet_data[nom_formatted][idx] += int(v) if v is not None else 0
                        except:
                            pass
                            
            # Add to labels_data
            for emp in sheet_employees:
                total = sum(sheet_data[emp])
                labels_data.append({
                    "nom": emp,
                    "total": total,
                    "period": period
                })
                
        if not labels_data:
            return jsonify({"error": "Aucune donnée d'employé trouvée dans le fichier Excel."}), 400
            
        # Sort labels by period, then by employee name
        labels_data.sort(key=lambda x: (x["period"], x["nom"]))
        
        pdf_bytes = generer_etiquettes_pdf(labels_data)
        
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=f"etiquettes_lait_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── Routes API ────────────────────────────────────────────────────────────
@app.route("/api/preview", methods=["POST"])
def preview():
    """Upload un fichier, retourne les données JSON pour le dashboard."""
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400
    f = request.files["file"]
    try:
        file_bytes = f.read()
        employees, dates, data = lire_excel_dynamique(file_bytes)

        # Calculs statistiques
        stats_employees = []
        total_global = 0
        for emp in employees:
            vals = data[emp]
            total = sum(vals)
            total_global += total
            stats_employees.append({
                "nom": emp,
                "total": total,
                "sans": len(dates) - total,
                "taux": round(total / len(dates) * 100, 1) if dates else 0,
                "valeurs": vals
            })

        stats_employees.sort(key=lambda x: x["total"], reverse=True)

        return jsonify({
            "employees": stats_employees,
            "dates": dates,
            "nb_dates": len(dates),
            "nb_employees": len(employees),
            "total_global": total_global,
            "taux_global": round(total_global / (len(employees) * len(dates)) * 100, 1) if employees and dates else 0
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/download", methods=["POST"])
def download():
    """Génère et retourne le fichier Excel ou CSV rapport."""
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400
    f = request.files["file"]
    try:
        file_bytes = f.read()
        is_csv = is_csv_bytes(file_bytes)
        
        # Load workbook
        wb = load_workbook_from_bytes(file_bytes)
        processed_wb = process_workbook_in_place(wb)
        
        if is_csv:
            # Export to CSV
            output = io.StringIO()
            writer = csv.writer(output, delimiter=';') # Use semicolon as standard for European Excel
            ws = processed_wb.active
            for row in ws.iter_rows(values_only=True):
                writer.writerow([val if val is not None else "" for val in row])
            
            csv_bytes = output.getvalue().encode('utf-8-sig') # UTF-8 with BOM for Excel compatibility
            return send_file(
                io.BytesIO(csv_bytes),
                mimetype="text/csv",
                as_attachment=True,
                download_name=f"rapport_total_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
            )
        else:
            # Export to Excel
            buf = io.BytesIO()
            processed_wb.save(buf)
            buf.seek(0)
            excel_bytes = buf.getvalue()
            return send_file(
                io.BytesIO(excel_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=f"rapport_total_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True, port=5000)
