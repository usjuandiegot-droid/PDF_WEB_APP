```python
import fitz
import pandas as pd
import re
import difflib
import os
import json
import traceback
import subprocess
import tempfile

from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image


# =========================
# APP
# =========================

app = Flask(__name__)
CORS(app)

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB


# =========================
# LOGS
# =========================

class JsonLogger:

    def log(self, level, event, message, extra=None):

        entry = {
            "timestamp": datetime.now(
                ZoneInfo("America/Bogota")
            ).isoformat(),
            "level": level,
            "event": event,
            "message": message
        }

        if extra:
            entry["extra"] = extra

        print(json.dumps(entry))

    def info(self, event, message, extra=None):
        self.log("info", event, message, extra)

    def error(self, event, message, extra=None):
        self.log("error", event, message, extra)


logger = JsonLogger()


# =========================
# BASE DIR
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

cie10_df = pd.read_excel(
    os.path.join(BASE_DIR, "CIE10.xlsx")
)

eps_df = pd.read_excel(
    os.path.join(BASE_DIR, "EPS.xlsx")
)

eps_df.columns = eps_df.columns.str.strip().str.upper()

eps_lista = eps_df["EPS"].dropna().tolist()

cie10_df["Codigo"] = (
    cie10_df["Codigo"]
    .astype(str)
    .str.strip()
    .str.upper()
)


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
# EXTRACCIÓN
# =========================

def extract_data(text):

    patient_name = None
    identification = None
    diagnosis = None
    date_of_attention = None
    order = None
    prorogation = None
    start_date = None
    end_date = None
    eps = None
    tipo_incapacidad = None
    grupo_servicio = None

    # =========================
    # PACIENTE
    # =========================

    m = re.search(
        r'PRESCRIPCIÓN DE INCAPACIDAD/LICENCIA DE MATERNIDAD\s*([^\n]*)\s*Identificación:',
        text
    )

    if m:
        patient_name = m.group(1).strip()

    # =========================
    # IDENTIFICACIÓN
    # =========================

    m = re.search(
        r'Identificación:\s*CC\s*(\d+)',
        text
    )

    if m:
        identification = m.group(1).strip()

    # =========================
    # DX
    # =========================

    m = re.search(
        r'\b([A-Z]\d{2,3}[A-Z]?)\b',
        text
    )

    if m:
        diagnosis = m.group(1).strip().upper()

    # =========================
    # EPS
    # =========================

    eps_lines = re.findall(
        r'EPS:\s*(.*)',
        text
    )

    if eps_lines:
        eps = homologar_eps(
            eps_lines[-1].strip()
        )

    # =========================
    # FECHA INICIO
    # =========================

    m = re.search(
        r'Fecha Inicio:\s*(\d{2}/\d{2}/\d{4})',
        text
    )

    if m:
        start_date = m.group(1)

    # =========================
    # FECHA FIN
    # =========================

    m = re.search(
        r'Fecha Fin:\s*(\d{2}/\d{2}/\d{4})',
        text
    )

    if m:
        end_date = m.group(1)

    # =========================
    # FECHA
    # =========================

    m = re.search(
        r'Orden:\s*\d+\n(\d{4}/\d{2}/\d{2})',
        text
    )

    if m:
        date_of_attention = m.group(1)

    # =========================
    # ORDEN
    # =========================

    m = re.search(
        r'Orden:\s*(\d+)',
        text
    )

    if m:
        order = m.group(1)

    # =========================
    # PRORROGA
    # =========================

    m = re.search(
        r'Prórroga:\s*(NO|SI)',
        text
    )

    if m:
        prorogation = m.group(1)

    # =========================
    # TIPO
    # =========================

    m = re.search(
        r'Tipo Incapacidad:\s*(.*)',
        text
    )

    if m:
        tipo_incapacidad = m.group(1).strip()

    # =========================
    # GRUPO
    # =========================

    m = re.search(
        r'Grupo Servicio:\s*(.*)',
        text
    )

    if m:
        grupo_servicio = m.group(1).strip()

    return {
        "Documento": identification,
        "Paciente": patient_name,
        "DX": diagnosis,
        "EPS": eps,
        "Fecha de Inicio": start_date,
        "Fecha de Fin": end_date,
        "FECHA": date_of_attention,
        "ORDEN": order,
        "PRORROGA": prorogation,
        "Tipo Incapacidad": tipo_incapacidad,
        "Grupo Servicio": grupo_servicio,
        "Dias": calculate_total_days(start_date, end_date)
    }


# =========================
# ENDPOINT
# =========================

@app.route("/procesar", methods=["POST"])

def procesar():

    try:

        files = request.files.getlist("files")

        now_co = datetime.now(
            ZoneInfo("America/Bogota")
        )

        logger.info(
            "inicio",
            "Procesamiento iniciado",
            {"total": len(files)}
        )

        all_data = []

        # =========================
        # PROCESAR PDFs
        # =========================

        for file in files:

            pdf_bytes = file.read()

            pdf = fitz.open(
                stream=pdf_bytes,
                filetype="pdf"
            )

            if len(pdf) == 0:
                continue

            text = pdf[0].get_text("text")

            data = extract_data(text)

            all_data.append(data)

            pdf.close()

        # =========================
        # DATAFRAME
        # =========================

        df = pd.DataFrame(all_data)

        df["CONCATENAR"] = (
            df["Documento"].astype(str)
            + " "
            + df["Paciente"].astype(str)
        )

        df["Total de Días"] = df["Dias"]

        df["Observacion"] = None

        df["Fecha envio dra"] = now_co.strftime(
            "%d/%m/%Y"
        )

        # =========================
        # HOMOLOGAR CIE10
        # =========================

        df["DX"] = (
            df["DX"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df = pd.merge(
            df,
            cie10_df,
            left_on="DX",
            right_on="Codigo",
            how="left"
        )

        df["DX"] = (
            df["DX"]
            + " - "
            + df["Nombre"].fillna("")
        )

        df.drop(
            columns=["Nombre", "Codigo"],
            inplace=True,
            errors="ignore"
        )

        # =========================
        # ORDEN COLUMNAS
        # =========================

        columnas = [
            "Fecha envio dra",
            "Documento",
            "Paciente",
            "CONCATENAR",
            "Fecha de Inicio",
            "Fecha de Fin",
            "Total de Días",
            "Dias",
            "EPS",
            "Observacion",
            "DX",
            "FECHA",
            "ORDEN",
            "PRORROGA",
            "Tipo Incapacidad",
            "Grupo Servicio"
        ]

        df = df.reindex(columns=columnas)

        # =========================
        # CARGAR PLANTILLA
        # =========================

        wb = load_workbook(
            os.path.join(BASE_DIR, "PLANTILLA.xlsx")
        )

        ws = wb["INFO"]

        # =========================
        # LIMPIAR SOLO DATOS
        # =========================

        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row)

        # =========================
        # INSERTAR DATA
        # =========================

        for r_idx, row in enumerate(
            dataframe_to_rows(
                df,
                index=False,
                header=False
            ),
            start=2
        ):

            for c_idx, value in enumerate(
                row,
                start=1
            ):

                ws.cell(
                    row=r_idx,
                    column=c_idx,
                    value=value
                )

        # =========================
        # INSERTAR FIRMAS
        # =========================
        firma_path = os.path.join(
            BASE_DIR,
            "firma.png"
        )

        if os.path.exists(firma_path):

            fila_base = 14
            salto = 31

            for i in range(len(df)):

                fila_firma = fila_base + (i * salto)

                firma = Image(firma_path)

                firma.width = 180
                firma.height = 80

                ws.add_image(
                    firma,
                    f"F{fila_firma}"
                )

        # =========================
        # GENERAR PDF
        # =========================

        with tempfile.TemporaryDirectory() as tmpdir:

            excel_path = os.path.join(
                tmpdir,
                "archivo.xlsx"
            )

            wb.save(excel_path)

            logger.info(
                "excel",
                "Excel temporal generado",
                {"path": excel_path}
            )

            # =========================
            # CONVERTIR A PDF
            # =========================

            subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                tmpdir,
                excel_path
            ], check=True)

            pdf_path = os.path.join(
                tmpdir,
                "archivo.pdf"
            )

            logger.info(
                "pdf",
                "PDF generado correctamente",
                {"path": pdf_path}
            )

            return send_file(
                pdf_path,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=(
                    f"Formato_Rechazos_Sura_"
                    f"{now_co.strftime('%d-%m-%Y')}.pdf"
                )
            )

    except Exception as e:

        logger.error(
            "error",
            "Fallo en procesamiento",
            {
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
```
