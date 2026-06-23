from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.utils import get_column_letter
import io
from datetime import datetime
import unicodedata
from copy import copy

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

    return header_row, emp_col, emp_rows, total_rows, sum_cols, existing_total_col

def process_workbook_in_place(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        header_row, emp_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
        
        if not emp_rows or not sum_cols:
            continue
            
        # Determine layout type
        name_counts = {}
        for r in emp_rows:
            val = ws.cell(row=r, column=emp_col).value
            if val is not None:
                name_counts[val] = name_counts.get(val, 0) + 1
                
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
                if val is not None:
                    nom_str = str(val).strip()
                    if nom_str not in emp_to_rows:
                        unique_emps.append(nom_str)
                        emp_to_rows[nom_str] = []
                    emp_to_rows[nom_str].append(r)
            
            rows_to_delete = []
            for emp_name in unique_emps:
                rows = emp_to_rows[emp_name]
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
            header_row, emp_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
            if not emp_rows or not sum_cols:
                continue

        # Find all existing total columns
        total_cols = []
        for c in range(emp_col + 1, ws.max_column + 1):
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
        
        # Populate formulas for employees
        for r in emp_rows:
            formula = f"=SUM({start_col_letter}{r}:{end_col_letter}{r})"
            cell = ws.cell(row=r, column=total_col, value=formula)
            copy_style(ws.cell(row=r, column=sum_cols[-1]), cell)
            
        # Populate formulas for summary rows
        if emp_rows:
            first_emp_row = emp_rows[0]
            last_emp_row = emp_rows[-1]
            for r in total_rows:
                formula = f"=SUM({total_col_letter}{first_emp_row}:{total_col_letter}{last_emp_row})"
                cell = ws.cell(row=r, column=total_col, value=formula)
                copy_style(ws.cell(row=r, column=sum_cols[-1]), cell)
                
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def lire_excel_dynamique(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row, emp_col, emp_rows, total_rows, sum_cols, existing_total_col = detect_structure(ws)
        if emp_rows and sum_cols:
            # Determine if List Layout (Layout B)
            name_counts = {}
            for r in emp_rows:
                val = ws.cell(row=r, column=emp_col).value
                if val is not None:
                    name_counts[val] = name_counts.get(val, 0) + 1
                    
            is_list_layout = False
            if len(emp_rows) > 0:
                max_occ = max(name_counts.values()) if name_counts else 0
                if max_occ > 1:
                    is_list_layout = True
                    
            if not is_list_layout:
                # Pivot Layout (Layout A)
                dates = []
                for c in sum_cols:
                    hdr_val = ws.cell(row=header_row, column=c).value
                    dates.append(str(hdr_val) if hdr_val is not None else "")
                    
                employees = []
                data = {}
                for r in emp_rows:
                    nom = ws.cell(row=r, column=emp_col).value
                    if nom is None:
                        continue
                    nom_str = str(nom).strip()
                    if nom_str not in data:
                        employees.append(nom_str)
                        data[nom_str] = []
                    
                    vals = []
                    for c in sum_cols:
                        v = ws.cell(row=r, column=c).value
                        try:
                            vals.append(int(v) if v is not None else 0)
                        except:
                            vals.append(0)
                    data[nom_str] = vals
                return employees, dates, data
            else:
                # List Layout (Layout B)
                # Aggregate values per employee
                dates = []
                for c in sum_cols:
                    hdr_val = ws.cell(row=header_row, column=c).value
                    dates.append(str(hdr_val) if hdr_val is not None else "")
                    
                employees_map = {} # nom_str -> list of sums for each sum_col
                for r in emp_rows:
                    nom = ws.cell(row=r, column=emp_col).value
                    if nom is None:
                        continue
                    nom_str = str(nom).strip()
                    if nom_str not in employees_map:
                        employees_map[nom_str] = [0] * len(sum_cols)
                        
                    for idx, c in enumerate(sum_cols):
                        v = ws.cell(row=r, column=c).value
                        try:
                            employees_map[nom_str][idx] += int(v) if v is not None else 0
                        except:
                            pass
                            
                # Sort unique employees
                employees = list(employees_map.keys())
                return employees, dates, employees_map
                
    return [], [], {}

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
    """Génère et retourne le fichier Excel rapport."""
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400
    f = request.files["file"]
    try:
        file_bytes = f.read()
        rapport_bytes = process_workbook_in_place(file_bytes)
        return send_file(
            io.BytesIO(rapport_bytes),
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
