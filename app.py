from flask import Flask, render_template, request, redirect, jsonify, send_file, session
import pandas as pd
import os
import json
import io
import base64

# For Excel styling
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

basedir = os.path.abspath(os.path.dirname(__file__))
templatedir = os.path.join(basedir, 'templates')
app = Flask(__name__, template_folder=templatedir)

# Secret key required to use client-side sessions securely
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "super-secret-key-change-in-production")

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

# Helper function to get DataFrame and Supervisors directly out of the Client-Side Session
def load_session_data():
    # Load Logs Dataframe
    logs_b64 = session.get('excel_data_b64', '')
    if logs_b64:
        try:
            excel_bytes = base64.b64decode(logs_b64)
            df = pd.read_excel(io.BytesIO(excel_bytes))
            # Ensure columns map perfectly
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            df = df[COLUMNS]
        except Exception:
            df = pd.DataFrame(columns=COLUMNS)
    else:
        df = pd.DataFrame(columns=COLUMNS)

    # Load Supervisors List
    supervisors = session.get('supervisors_list', [])
    return df, supervisors

# Helper function to process styling on a raw bytes engine and return clean base64 string
def style_and_save_session_excel(df):
    output = io.BytesIO()
    
    # Initial save to bytes buffer
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11, color="000000")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = data_font
            cell.border = thin_border

            if col in [1, 2, 5, 6, 7, 8]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[1].height = 28
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 20

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if len(val) > max_len:
                max_len = len(val)
        clamped_width = min(max(max_len + 4, 12), 45)
        ws.column_dimensions[col_letter].width = clamped_width

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    
    # Store dynamic updates straight back to encrypted cookie memory
    session['excel_data_b64'] = base64.b64encode(final_output.getvalue()).decode('utf-8')

@app.route("/")
def home():
    session.permanent = True  
    app.permanent_session_lifetime = 31536000 # 1 year persistence window

    df, supervisors = load_session_data()
    
    # Run structural styles if session is blank or empty
    if 'excel_data_b64' not in session:
        style_and_save_session_excel(df)

    for col in ["Start Date", "End Date"]:
        if col in df.columns:
            df[col] = df[col].astype(str).replace("NaT", "").replace("nan", "")

    logs = df.to_dict(orient="records")
    logs_with_index = list(enumerate(logs))
    
    return render_template("index.html", logs=logs_with_index, supervisors=supervisors)

@app.route("/save", methods=["POST"])
def save():
    df, _ = load_session_data()
    
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

    if row_index_str != "":
        idx = int(row_index_str)
        if 0 <= idx < len(df):
            for col in COLUMNS:
                df.at[idx, col] = updated_row[col]
    else:
        df = pd.concat([df, pd.DataFrame([updated_row])], ignore_index=True)

    style_and_save_session_excel(df)
    return redirect("/")

@app.route("/delete/row/<int:index>", methods=["POST"])
def delete_row(index):
    try:
        df, _ = load_session_data()
        if 0 <= index < len(df):
            df = df.drop(index).reset_index(drop=True)
            style_and_save_session_excel(df)
    except Exception as e:
        print(f"Error deleting row: {e}")
    return redirect("/")

@app.route("/add_supervisor", methods=["POST"])
def addsupervisor():
    _, sups = load_session_data()
    name = request.form.get("name", "").strip()
    if name and name not in sups:
        sups.append(name)
        session['supervisors_list'] = sups
    return redirect("/")

@app.route("/delete/supervisor/<string:name>", methods=["POST"])
def delete_supervisor(name):
    _, sups = load_session_data()
    if name in sups:
        sups.remove(name)
        session['supervisors_list'] = sups
    return redirect("/")

@app.route("/get_last_entry")
def getlastentry():
    df, _ = load_session_data()
    if not df.empty:
        last_row = df.iloc[-1].fillna("").to_dict()
        return jsonify(last_row)
    return jsonify({})

@app.route("/open_excel")
def openexcel():
    try:
        logs_b64 = session.get('excel_data_b64', '')
        if logs_b64:
            excel_bytes = base64.b64decode(logs_b64)
            return send_file(
                io.BytesIO(excel_bytes),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="internshiplog.xlsx"
            )
        else:
            # Fallback to building an empty framework sheet instantly
            df = pd.DataFrame(columns=COLUMNS)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)
            return send_file(
                output,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="internshiplog.xlsx"
            )
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
