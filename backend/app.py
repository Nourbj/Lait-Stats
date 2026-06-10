from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import io
import os
import json
from datetime import datetime

app = Flask(__name__)
CORS(app)

# ── helpers ────────────────────────────────────────────────────────────────
def border_thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def lire_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    max_col = ws.max_column
    max_row = ws.max_row

    dates = []
    for col in range(2, max_col + 1):
        val = ws.cell(row=1, column=col).value
        if val is not None:
            dates.append(str(val))

    employees = []
    data = {}
    for row in range(2, max_row + 1):
        nom = ws.cell(row=row, column=1).value
        if not nom:
            continue
        nom = str(nom).strip()
        employees.append(nom)
        valeurs = []
        for col in range(2, 2 + len(dates)):
            v = ws.cell(row=row, column=col).value
            try:
                valeurs.append(int(v) if v is not None else 0)
            except:
                valeurs.append(0)
        data[nom] = valeurs

    return employees, dates, data

def generer_rapport(employees, dates, data):
    BLEU  = "1A3A5C"
    VERT  = "0EA5A0"
    BLANC = "FFFFFF"
    GRIS  = "F7F8FA"

    wb = openpyxl.Workbook()
    nb_dates = len(dates)

    # ── Feuille 1: Détail ─────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Détail Journalier"

    header_fill = PatternFill("solid", start_color=BLEU)
    header_font = Font(bold=True, color=BLANC, name="Calibri", size=10)
    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")

    ws1.cell(row=1, column=1, value="Employé").font      = header_font
    ws1.cell(row=1, column=1).fill      = header_fill
    ws1.cell(row=1, column=1).alignment = left
    ws1.cell(row=1, column=1).border    = border_thin()

    for ci, d in enumerate(dates, start=2):
        c = ws1.cell(row=1, column=ci, value=d)
        c.font = header_font; c.fill = header_fill
        c.alignment = center; c.border = border_thin()

    total_col = 2 + nb_dates
    c = ws1.cell(row=1, column=total_col, value="TOTAL")
    c.font = Font(bold=True, color=BLANC, name="Calibri", size=10)
    c.fill = PatternFill("solid", start_color=VERT)
    c.alignment = center; c.border = border_thin()

    fill_pair   = PatternFill("solid", start_color=GRIS)
    fill_impair = PatternFill("solid", start_color=BLANC)
    fill_1      = PatternFill("solid", start_color="D1FAF8")
    font_normal = Font(name="Calibri", size=10)
    font_bold   = Font(name="Calibri", size=10, bold=True)

    for ri, emp in enumerate(employees, start=2):
        alt = fill_pair if ri % 2 == 0 else fill_impair
        c = ws1.cell(row=ri, column=1, value=emp)
        c.font = font_bold; c.fill = alt; c.alignment = left; c.border = border_thin()
        vals = data[emp]
        for ci, v in enumerate(vals, start=2):
            cell = ws1.cell(row=ri, column=ci, value=v)
            cell.font = font_normal; cell.alignment = center; cell.border = border_thin()
            cell.fill = fill_1 if v == 1 else fill_impair

        tl = ws1.cell(row=ri, column=total_col)
        tl.value = f"=SUM({get_column_letter(2)}{ri}:{get_column_letter(1+nb_dates)}{ri})"
        tl.font = Font(name="Calibri", size=10, bold=True, color=BLANC)
        tl.fill = PatternFill("solid", start_color=VERT)
        tl.alignment = center; tl.border = border_thin()

    last_row = 2 + len(employees)
    c = ws1.cell(row=last_row, column=1, value="TOTAL")
    c.font = Font(bold=True, color=BLANC, name="Calibri"); c.fill = header_fill
    c.alignment = left; c.border = border_thin()
    for ci in range(2, total_col + 1):
        let = get_column_letter(ci)
        c = ws1.cell(row=last_row, column=ci)
        c.value = f"=SUM({let}2:{let}{last_row-1})"
        c.font = Font(bold=True, color=BLANC, name="Calibri")
        c.fill = header_fill; c.alignment = center; c.border = border_thin()

    ws1.column_dimensions["A"].width = 24
    for ci in range(2, total_col):
        ws1.column_dimensions[get_column_letter(ci)].width = 7
    ws1.column_dimensions[get_column_letter(total_col)].width = 10
    ws1.row_dimensions[1].height = 28
    ws1.freeze_panes = "B2"

    # ── Feuille 2: Récapitulatif ──────────────────────────────────────────
    ws2 = wb.create_sheet("Récapitulatif")
    ws2.merge_cells("A1:D1")
    t = ws2["A1"]
    t.value = f"Récapitulatif — {datetime.now().strftime('%d/%m/%Y')}"
    t.font  = Font(bold=True, name="Calibri", size=13, color=BLANC)
    t.fill  = PatternFill("solid", start_color=BLEU)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 30

    hdrs = ["Employé","Paquets pris","Jours sans lait","Taux de prise"]
    hf = [BLEU, VERT, "E07B39", BLEU]
    for ci, (h, fc) in enumerate(zip(hdrs, hf), start=1):
        c = ws2.cell(row=2, column=ci, value=h)
        c.font = Font(bold=True, color=BLANC, name="Calibri", size=10)
        c.fill = PatternFill("solid", start_color=fc); c.alignment = center; c.border = border_thin()

    tl = get_column_letter(total_col)
    for ri, emp in enumerate(employees, start=3):
        src = ri - 3 + 2
        alt = fill_pair if ri % 2 == 1 else fill_impair
        c = ws2.cell(row=ri, column=1, value=emp)
        c.font = font_bold; c.fill = alt; c.alignment = left; c.border = border_thin()
        c2 = ws2.cell(row=ri, column=2)
        c2.value = f"='Détail Journalier'!{tl}{src}"; c2.font = Font(name="Calibri", size=10, bold=True, color="0EA5A0")
        c2.fill = PatternFill("solid", start_color="D1FAF8"); c2.alignment = center; c2.border = border_thin()
        c3 = ws2.cell(row=ri, column=3)
        c3.value = f"={nb_dates}-'Détail Journalier'!{tl}{src}"
        c3.font = Font(name="Calibri", size=10, bold=True, color="A04000")
        c3.fill = PatternFill("solid", start_color="FDEBD0"); c3.alignment = center; c3.border = border_thin()
        c4 = ws2.cell(row=ri, column=4)
        c4.value = f"='Détail Journalier'!{tl}{src}/{nb_dates}"; c4.number_format = "0.0%"
        c4.font = Font(name="Calibri", size=10, bold=True); c4.fill = alt; c4.alignment = center; c4.border = border_thin()

    lr = 2 + len(employees) + 1
    c = ws2.cell(row=lr, column=1, value="TOTAL")
    c.font = Font(bold=True, color=BLANC, name="Calibri"); c.fill = header_fill
    c.alignment = left; c.border = border_thin()
    for ci in range(2, 5):
        let = get_column_letter(ci)
        c = ws2.cell(row=lr, column=ci)
        c.value = f"=SUM({let}3:{let}{lr-1})" if ci < 4 else f"=AVERAGE(D3:D{lr-1})"
        c.number_format = "0.0%" if ci == 4 else "General"
        c.font = Font(bold=True, color=BLANC, name="Calibri"); c.fill = header_fill
        c.alignment = center; c.border = border_thin()

    for col, w in zip("ABCD", [24, 18, 18, 16]):
        ws2.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── Routes API ────────────────────────────────────────────────────────────
@app.route("/api/preview", methods=["POST"])
def preview():
    """Upload un fichier, retourne les données JSON pour le dashboard."""
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400
    f = request.files["file"]
    try:
        file_bytes = f.read()
        employees, dates, data = lire_excel(file_bytes)

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
        employees, dates, data = lire_excel(file_bytes)
        rapport_bytes = generer_rapport(employees, dates, data)
        return send_file(
            io.BytesIO(rapport_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"rapport_lait_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True, port=5000)
