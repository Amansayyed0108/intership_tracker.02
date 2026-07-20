from flask import Flask, render_template, request, redirect, jsonify
import pandas as pd
import os
import json
import subprocess
import platform

# For Excel styling
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

basedir = os.path.abspath(os.path.dirname(__file__))
templatedir = os.path.join(basedir, 'templates')
app = Flask(__name__, template_folder=templatedir)

EXCELFILE = os.path.join(basedir, "internshiplog.xlsx")
SUPERVISORSFILE = os.path.join(basedir, "supervisors.json")

COLUMNS = [
    "Start Date",
    "End Date",
    "Project Task",
    "Description",
    "Tool",
    "Hours",
    "Status",
    "Supervisor"
]

# Helper function to apply beautiful professional styling to the Excel file
def styleexcel():
    if not os.path.exists(EXCELFILE):
        return

    wb = load_workbook(EXCELFILE)
    ws = wb.active

    # Theme Colors (Professional Dark Navy Blue / Clean White text)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="000000")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Style Header Row
    for col_idx in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Style Data Rows
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border

            # Center dates, status, hours and supervisor
            if col in [1, 2, 5, 6, 7, 8]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Set custom row heights
    ws.row_dimensions[1].height = 28  # Thicker header row
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 20  # Comfortable reading height for data

    # Auto-adjust column widths cleanly based on data length
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if len(val) > max_len:
                max_len = len(val)
        # Prevent columns from becoming ridiculously wide for long descriptions
        clamped_width = min(max(max_len + 4, 12), 45)
        ws.column_dimensions[col_letter].width = clamped_width

    wb.save(EXCELFILE)

# Ensure Excel file exists with correct headers
def initexcel():
    if not os.path.exists(EXCELFILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(EXCELFILE, index=False)
        styleexcel()
    else:
        try:
            df = pd.read_excel(EXCELFILE)
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            df.to_excel(EXCELFILE, index=False)
            styleexcel()
        except Exception:
            df = pd.DataFrame(columns=COLUMNS)
            df.to_excel(EXCELFILE, index=False)
            styleexcel()

initexcel()

def loadsupervisors():
    if os.path.exists(SUPERVISORSFILE):
        try:
            with open(SUPERVISORSFILE, 'r') as f:
                return json.load(f)
        except Exception:
            return []
    return []

def savesupervisors(suplist):
    with open(SUPERVISORSFILE, 'w') as f:
        json.dump(suplist, f)

@app.route("/")
def home():
    df = pd.read_excel(EXCELFILE)
    for col in ["Start Date", "End Date"]:
        if col in df.columns:
            df[col] = df[col].astype(str).replace("NaT", "").replace("nan", "")

    logs = df.to_dict(orient="records")
    # Zip data with its DataFrame raw index number to let us target it securely
    logs_with_index = list(enumerate(logs))
    
    supervisors = loadsupervisors()
    return render_template("index.html", logs=logs_with_index, supervisors=supervisors)

@app.route("/save", methods=["POST"])
def save():
    row_index_str = request.form.get("row_index", "")
    start_date = request.form.get("start_date", "")
    end_date = request.form.get("end_date", "")
    task = request.form.get("task", "")
    description = request.form.get("description", "")
    tool = request.form.get("tool", "")
    hours = request.form.get("hours", "0")
    minutes = request.form.get("minutes", "00")
    status = request.form.get("status", "")
    supervisor = request.form.get("supervisor", "")

    formatted_time = f"{int(hours):02d}:{int(minutes):02d}"

    updated_row = {
        "Start Date": start_date,
        "End Date": end_date,
        "Project Task": task,
        "Description": description,
        "Tool": tool,
        "Hours": formatted_time,
        "Status": status,
        "Supervisor": supervisor
    }

    df = pd.read_excel(EXCELFILE)

    # Check if this request is updating an existing entry
    if row_index_str != "":
        idx = int(row_index_str)
        if 0 <= idx < len(df):
            # Replace row values directly at targeted index location
            for col in COLUMNS:
                df.at[idx, col] = updated_row[col]
    else:
        # Append new log row entry normally
        df = pd.concat([df, pd.DataFrame([updated_row])], ignore_index=True)

    df.to_excel(EXCELFILE, index=False)
    styleexcel()
    return redirect("/")

@app.route("/delete/row/<int:index>", methods=["POST"])
def delete_row(index):
    try:
        df = pd.read_excel(EXCELFILE)
        if 0 <= index < len(df):
            df = df.drop(index).reset_index(drop=True)
            df.to_excel(EXCELFILE, index=False)
            styleexcel()
    except Exception as e:
        print(f"Error deleting row: {e}")
    return redirect("/")

@app.route("/add_supervisor", methods=["POST"])
def addsupervisor():
    name = request.form.get("name", "").strip()
    if name:
        sups = loadsupervisors()
        if name not in sups:
            sups.append(name)
            savesupervisors(sups)
    return redirect("/")

@app.route("/delete/supervisor/<string:name>", methods=["POST"])
def delete_supervisor(name):
    sups = loadsupervisors()
    if name in sups:
        sups.remove(name)
        savesupervisors(sups)
    return redirect("/")

@app.route("/get_last_entry")
def getlastentry():
    df = pd.read_excel(EXCELFILE)
    if not df.empty:
        last_row = df.iloc[-1].fillna("").to_dict()
        return jsonify(last_row)
    return jsonify({})

@app.route("/open_excel")
def openexcel():
    try:
        if platform.system() == "Windows":
            os.startfile(EXCELFILE)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", EXCELFILE])
        else:  # Linux
            subprocess.run(["xdg-open", EXCELFILE])
        return jsonify({"status": "success"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, port=8080)