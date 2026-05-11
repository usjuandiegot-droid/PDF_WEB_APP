import fitz
import pandas as pd
import re
import difflib
import os
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

cie10_df = pd.read_excel(os.path.join(BASE_DIR, "CIE10.xlsx"))
eps_df = pd.read_excel(os.path.join(BASE_DIR, "EPS.xlsx"))

eps_df.columns = eps_df.columns.str.strip().str.upper()
eps_lista = eps_df["EPS"].dropna().tolist()

cie10_df["Codigo"] = cie10_df["Codigo"].astype(str).str.strip().str.upper()


# =========================
# EPS HOMOLOGACIÓN (MEJORADA PERO COMPATIBLE)
# =========================

def homologar_eps(eps_extraida):

    if not eps_extraida:
        return None

    eps_extraida = eps_extraida.strip().upper()

    matches = difflib.get_close_matches(
        eps_extraida,
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
# TU EXTRACTOR ORIGINAL (SIN ROMPERLO)
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

    # PACIENTE
    m = re.search(
        r'PRESCRIPCIÓN DE INCAPACIDAD/LICENCIA DE MATERNIDAD\s*([^\n]*)\s*Identificación:',
        text
    )
    if m:
        patient_name = m.group(1).strip()

    # IDENTIFICACIÓN
    m = re.search(r'Identificación:\s*CC\s*(\d+)', text)
    if m:
        identification = m.group(1).strip()

    # DX
    m = re.search(r'\b([A-Z]\d{2,3}[A-Z]?)\b', text)
    if m:
        diagnosis = m.group(1).strip().upper()

    # EPS (MEJORADO SOLO AQUÍ)
    eps_lines = re.findall(r'EPS:\s*(.*)', text)
    if eps_lines:
        eps = homologar_eps(eps_lines[-1])

    # FECHAS
    m = re.search(r'Fecha Inicio:\s*(\d{2}/\d{2}/\d{4})', text)
    if m:
        start_date = m.group(1)

    m = re.search(r'Fecha Fin:\s*(\d{2}/\d{2}/\d{4})', text)
    if m:
        end_date = m.group(1)

    # ORDEN / FECHA ATENCIÓN
    m = re.search(r'Orden:\s*\d+\n(\d{4}/\d{2}/\d{2})', text)
    if m:
        date_of_attention = m.group(1)

    m = re.search(r'Orden:\s*(\d+)', text)
    if m:
        order = m.group(1)

    # PRÓRROGA
    m = re.search(r'Prórroga:\s*(NO|SI)', text)
    if m:
        prorogation = m.group(1)

    # TIPO
    m = re.search(r'Tipo Incapacidad:\s*(.*)', text)
    if m:
        tipo_incapacidad = m.group(1).strip()

    # GRUPO
    m = re.search(r'Grupo Servicio:\s*(.*)', text)
    if m:
        grupo_servicio = m.group(1).strip()

    return {
        "Documento": identification,
        "Paciente": patient_name,
        "DX": diagnosis,
        "EPS": eps,
        "Fecha Inicio": start_date,
        "Fecha Fin": end_date,
        "FECHA": date_of_attention,
        "ORDEN": order,
        "PRORROGA": prorogation,
        "Tipo Incapacidad": tipo_incapacidad,
        "Grupo Servicio": grupo_servicio,
        "Dias": None
    }


# =========================
# ENDPOINT
# =========================

@app.route("/procesar", methods=["POST"])

def procesar():

    try:

        files = request.files.getlist("files")

        all_data = []

        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

            for file in files:

                pdf_bytes = file.read()

                pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

                text = pdf[0].get_text("text")

                data = extract_data(text)
                all_data.append(data)

                # guardar PDF original
                zip_file.writestr(file.filename, pdf_bytes)

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

            df["DX"] = df["DX"] + " - " + df["Nombre"]

            df.drop(columns=["Nombre", "Codigo"], inplace=True, errors="ignore")

            df.insert(0, "Fecha envio dra", datetime.now().strftime("%d/%m/%Y"))

            # =========================
            # EXCEL
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

            excel_buffer = BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)

            zip_file.writestr(
                f"Formato_Rechazos_Sura_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
                excel_buffer.read()
            )

        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name="resultado.zip"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run()
