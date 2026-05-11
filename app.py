import fitz
import pandas as pd
import re
import difflib
import os
import json
import traceback
import zipfile

from datetime import datetime
from io import BytesIO

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# =========================
# APP
# =========================

app = Flask(__name__)
CORS(app)


# =========================
# LOGS
# =========================

class JsonLogger:

    def _log(self, level, event, message, extra=None):

        log_entry = {
            "timestamp": datetime.utcnow().isoformat(),
            "level": level,
            "event": event,
            "message": message
        }

        if extra:
            log_entry["extra"] = extra

        print(json.dumps(log_entry))

    def info(self, event, message, extra=None):
        self._log("info", event, message, extra)

    def error(self, event, message, extra=None):
        self._log("error", event, message, extra)


logger = JsonLogger()


# =========================
# ARCHIVOS BASE
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

cie10_df = pd.read_excel(os.path.join(BASE_DIR, "CIE10.xlsx"))
eps_df = pd.read_excel(os.path.join(BASE_DIR, "EPS.xlsx"))
wb_template = load_workbook(os.path.join(BASE_DIR, "PLANTILLA.xlsx"))

eps_df.columns = eps_df.columns.str.strip().str.upper()
eps_lista = eps_df["EPS"].dropna().tolist()

cie10_df["Codigo"] = cie10_df["Codigo"].astype(str).str.strip().str.upper()


# =========================
# FUNCIONES
# =========================

def calculate_total_days(start, end):

    if not start or not end:
        return None

    try:
        return (
            datetime.strptime(end, "%d/%m/%Y")
            - datetime.strptime(start, "%d/%m/%Y")
        ).days + 1

    except:
        return None


def homologar_eps(eps_extraida):

    if not eps_extraida:
        return None

    matches = difflib.get_close_matches(
        eps_extraida.upper(),
        [e.upper() for e in eps_lista],
        n=1,
        cutoff=0.5
    )

    if matches:
        for eps in eps_lista:
            if eps.upper() == matches[0]:
                return eps

    return eps_extraida


# =========================
# ENDPOINT
# =========================

@app.route("/procesar", methods=["POST"])

def procesar():

    try:

        files = request.files.getlist("files")

        logger.info("inicio", "Procesamiento iniciado", {"total": len(files)})

        all_data = []

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

            for file in files:

                file_bytes = file.read()

                pdf = fitz.open(stream=file_bytes, filetype="pdf")

                if len(pdf) == 0:
                    continue

                text = pdf[0].get_text("text")

                doc = re.search(r'Identificación:\s*CC\s*(\d+)', text)
                name = re.search(r'PRESCRIPCIÓN.*?\n(.*?)\nIdentificación:', text)
                dx = re.search(r'\b([A-Z]\d{2,3}[A-Z]?)\b', text)

                eps = re.findall(r'EPS:\s*(.*)', text)
                start = re.search(r'Fecha Inicio:\s*(\d{2}/\d{2}/\d{4})', text)
                end = re.search(r'Fecha Fin:\s*(\d{2}/\d{2}/\d{4})', text)

                all_data.append({
                    "Documento": doc.group(1) if doc else None,
                    "Paciente": name.group(1).strip() if name else None,
                    "DX": dx.group(1).strip().upper() if dx else None,
                    "EPS": homologar_eps(eps[-1].strip()) if eps else None,
                    "Inicio": start.group(1) if start else None,
                    "Fin": end.group(1) if end else None,
                    "Dias": calculate_total_days(
                        start.group(1) if start else None,
                        end.group(1) if end else None
                    )
                })

                # 👉 agregar PDF al ZIP
                zip_file.writestr(file.filename, file_bytes)

                pdf.close()

        # =========================
        # DATAFRAME
        # =========================

        df = pd.DataFrame(all_data)

        df["DX"] = df["DX"].astype(str).str.strip().str.upper()

        df = pd.merge(
            df,
            cie10_df,
            left_on="DX",
            right_on="Codigo",
            how="left"
        )

        df["DX"] = (
            df["DX"].fillna("") + " - " + df["Nombre"].fillna("")
        )

        df.drop(columns=["Nombre", "Codigo"], inplace=True, errors="ignore")

        df.insert(0, "Fecha Generación", datetime.now().strftime("%d/%m/%Y"))

        # =========================
        # EXCEL PLANTILLA
        # =========================

        wb = load_workbook(os.path.join(BASE_DIR, "PLANTILLA.xlsx"))
        ws = wb["INFO"]

        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row)

        for r_idx, row in enumerate(
            dataframe_to_rows(df, index=False, header=False),
            start=2
        ):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # =========================
        # AGREGAR EXCEL AL ZIP
        # =========================

        excel_buffer = BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)

        zip_file.writestr(
            f"Formato_Rechazos_Sura_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
            excel_buffer.read()
        )

        zip_buffer.seek(0)

        logger.info("fin", "ZIP generado correctamente", {"registros": len(df)})

        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name="resultado.zip"
        )

    except Exception as e:

        logger.error(
            "error",
            "Error en procesamiento",
            {
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )

        return jsonify({"error": str(e)}), 500


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run()
