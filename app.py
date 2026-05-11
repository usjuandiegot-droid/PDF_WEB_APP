import fitz
import pandas as pd
import re
import difflib
import os
import json
import traceback

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
# LOGGER ESTRUCTURADO
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

    def warning(self, event, message, extra=None):
        self._log("warning", event, message, extra)


logger = JsonLogger()


# =========================
# BASE DIR (IMPORTANTE RENDER)
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


# =========================
# CARGA ARCHIVOS
# =========================

cie10_df = pd.read_excel(os.path.join(BASE_DIR, "CIE10.xlsx"))
eps_df = pd.read_excel(os.path.join(BASE_DIR, "EPS.xlsx"))
wb_template = load_workbook(os.path.join(BASE_DIR, "PLANTILLA.xlsx"))

eps_df.columns = eps_df.columns.str.strip().str.upper()
eps_lista = eps_df["EPS"].tolist()


# =========================
# FUNCIONES
# =========================

def calculate_total_days(start_date_str, end_date_str):

    if start_date_str is None or end_date_str is None:
        return None

    try:
        start_date = datetime.strptime(start_date_str, "%d/%m/%Y")
        end_date = datetime.strptime(end_date_str, "%d/%m/%Y")

        return (end_date - start_date).days + 1

    except Exception:
        return None


def homologar_eps(eps_extraida):

    if not eps_extraida:
        return None

    coincidencias = difflib.get_close_matches(
        eps_extraida.upper(),
        [x.upper() for x in eps_lista],
        n=1,
        cutoff=0.5
    )

    if coincidencias:

        for eps in eps_lista:
            if eps.upper() == coincidencias[0]:
                return eps

    return eps_extraida


# =========================
# RUTA PRINCIPAL
# =========================

@app.route('/procesar', methods=['POST'])

def procesar_pdfs():

    request_id = str(datetime.utcnow().timestamp())

    try:

        files = request.files.getlist("files")

        logger.info(
            event="procesamiento_inicio",
            message="Inicio de procesamiento",
            extra={"request_id": request_id, "total_archivos": len(files)}
        )

        all_data = []

        for file in files:

            try:

                pdf_document = fitz.open(
                    stream=file.read(),
                    filetype="pdf"
                )

                if len(pdf_document) == 0:
                    logger.warning(
                        "pdf_vacio",
                        "PDF sin páginas",
                        {"file": file.filename}
                    )
                    continue

                text = pdf_document[0].get_text("text")

                # =========================
                # VARIABLES
                # =========================

                patient_name = None
                identification = None
                diagnosis = None
                start_date = None
                end_date = None
                eps = None

                # =========================
                # EXTRACCIÓN
                # =========================

                patient_match = re.search(
                    r'PRESCRIPCIÓN DE INCAPACIDAD/LICENCIA DE MATERNIDAD\s*([^\n]*)\s*Identificación:',
                    text
                )
                if patient_match:
                    patient_name = patient_match.group(1).strip()

                id_match = re.search(
                    r'Identificación:\s*CC\s*(\d+)',
                    text
                )
                if id_match:
                    identification = id_match.group(1)

                dx_match = re.search(r'\b([A-Z]\d{2,3}[A-Z]?)\b', text)
                if dx_match:
                    diagnosis = dx_match.group(1)

                eps_match = re.findall(r'EPS:\s*(.*)', text)
                if eps_match:
                    eps = homologar_eps(eps_match[-1].strip())

                start_match = re.search(r'Fecha Inicio:\s*(\d{2}/\d{2}/\d{4})', text)
                if start_match:
                    start_date = start_match.group(1)

                end_match = re.search(r'Fecha Fin:\s*(\d{2}/\d{2}/\d{4})', text)
                if end_match:
                    end_date = end_match.group(1)

                total_days = calculate_total_days(start_date, end_date)

                all_data.append({

                    "Documento": identification,
                    "Paciente": patient_name,
                    "DX": diagnosis,
                    "EPS": eps,
                    "Fecha Inicio": start_date,
                    "Fecha Fin": end_date,
                    "Total Días": total_days

                })

                pdf_document.close()

                logger.info(
                    event="pdf_procesado",
                    message="PDF procesado correctamente",
                    extra={"file": file.filename}
                )

            except Exception as e:

                logger.error(
                    event="error_pdf_individual",
                    message="Error procesando PDF",
                    extra={
                        "file": file.filename,
                        "error": str(e),
                        "trace": traceback.format_exc()
                    }
                )

        # =========================
        # DATAFRAME
        # =========================

        df = pd.DataFrame(all_data)

        df = pd.merge(
            df,
            cie10_df,
            left_on="DX",
            right_on="Codigo",
            how="left"
        )

        df["DX"] = df["DX"].fillna("") + " " + df["Nombre"].fillna("")
        df.drop(columns=["Nombre", "Codigo"], inplace=True, errors="ignore")

        df.insert(
            0,
            "Fecha Generación",
            datetime.now().strftime("%d/%m/%Y")
        )

        # =========================
        # PLANTILLA
        # =========================

        wb = load_workbook(os.path.join(BASE_DIR, "PLANTILLA.xlsx"))

        if "INFO" not in wb.sheetnames:
            return jsonify({"error": "Hoja INFO no existe"}), 500

        ws = wb["INFO"]

        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row)

        for r_idx, row in enumerate(
            dataframe_to_rows(df, index=False, header=False),
            start=2
        ):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        logger.info(
            event="procesamiento_finalizado",
            message="Archivo generado correctamente",
            extra={"request_id": request_id, "registros": len(df)}
        )

        return send_file(
            output,
            download_name=f"Formato Rechazos Sura {datetime.now().strftime('%d-%m-%Y')}.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        logger.error(
            event="error_general",
            message="Error en endpoint /procesar",
            extra={
                "error": str(e),
                "trace": traceback.format_exc()
            }
        )

        return jsonify({
            "error": str(e)
        }), 500


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run()
