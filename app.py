import fitz
import pandas as pd
import re
import difflib

from datetime import datetime
from io import BytesIO

from flask import Flask, request, send_file
from flask_cors import CORS

from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# =========================
# HABILITAR CORS
# =========================

CORS(app)

# =========================
# CARGAR ARCHIVOS
# =========================

# CIE10
cie10_df = pd.read_excel("CIE10.xlsx")

# EPS parametrizadas
eps_df = pd.read_excel("EPS.xlsx")

# Limpiar nombres columnas
eps_df.columns = (
    eps_df.columns
    .str.strip()
    .str.upper()
)

# Lista EPS
eps_lista = eps_df["EPS"].tolist()

# =========================
# FUNCIONES
# =========================

# Calcular días
def calculate_total_days(start_date_str, end_date_str):

    if start_date_str is None or end_date_str is None:
        return None

    try:

        start_date = datetime.strptime(
            start_date_str,
            "%d/%m/%Y"
        )

        end_date = datetime.strptime(
            end_date_str,
            "%d/%m/%Y"
        )

        total_days = (
            end_date - start_date
        ).days + 1

        return total_days

    except ValueError:

        return None

# Homologar EPS
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

    try:

        files = request.files.getlist("files")

        all_data = []

        for file in files:

            # Abrir PDF desde memoria
            pdf_document = fitz.open(
                stream=file.read(),
                filetype="pdf"
            )

            # Primera página
            first_page = pdf_document[0]

            text = first_page.get_text("text")

            # =========================
            # VARIABLES
            # =========================

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

            patient_name_match = re.search(
                r'PRESCRIPCIÓN DE INCAPACIDAD/LICENCIA DE MATERNIDAD\s*([^\n]*)\s*Identificación:',
                text
            )

            if patient_name_match:

                patient_name = patient_name_match.group(1).strip()

            # =========================
            # DOCUMENTO
            # =========================

            identification_match = re.search(
                r'Identificación:\s*CC\s*(\d+)',
                text
            )

            if identification_match:

                identification = identification_match.group(1).strip()

            # =========================
            # DIAGNÓSTICO
            # =========================

            diagnosis_code_match = re.search(
                r'\b([A-Z]\d{2,3}[A-Z]?)\b',
                text
            )

            if diagnosis_code_match:

                diagnosis = diagnosis_code_match.group(1).strip()

            # =========================
            # EPS
            # =========================

            eps_lines = re.findall(
                r'EPS:\s*(.*)',
                text
            )

            if eps_lines:

                eps_original = eps_lines[-1].strip()

                eps = homologar_eps(
                    eps_original
                )

            # =========================
            # FECHA INICIO
            # =========================

            start_date_match = re.search(
                r'Fecha Inicio:\s*(\d{2}/\d{2}/\d{4})',
                text
            )

            if start_date_match:

                start_date = start_date_match.group(1).strip()

            # =========================
            # FECHA FIN
            # =========================

            end_date_match = re.search(
                r'Fecha Fin:\s*(\d{2}/\d{2}/\d{4})',
                text
            )

            if end_date_match:

                end_date = end_date_match.group(1).strip()

            # =========================
            # FECHA ATENCIÓN
            # =========================

            date_of_attention_match = re.search(
                r'Orden:\s*\d+\n(\d{4}/\d{2}/\d{2})',
                text
            )

            if date_of_attention_match:

                date_of_attention = date_of_attention_match.group(1).strip()

            # =========================
            # ORDEN
            # =========================

            order_match = re.search(
                r'Orden:\s*(\d+)',
                text
            )

            if order_match:

                order = order_match.group(1).strip()

            # =========================
            # PRÓRROGA
            # =========================

            prorogation_match = re.search(
                r'Prórroga:\s*(NO|SI)',
                text
            )

            if prorogation_match:

                prorogation = prorogation_match.group(1).strip()

            # =========================
            # TIPO INCAPACIDAD
            # =========================

            tipo_incapacidad_match = re.search(
                r'Tipo Incapacidad:\s*(.*)',
                text
            )

            if tipo_incapacidad_match:

                tipo_incapacidad = tipo_incapacidad_match.group(1).strip()

            # =========================
            # GRUPO SERVICIO
            # =========================

            grupo_servicio_match = re.search(
                r'Grupo Servicio:\s*(.*)',
                text
            )

            if grupo_servicio_match:

                grupo_servicio = grupo_servicio_match.group(1).strip()

            # =========================
            # TOTAL DÍAS
            # =========================

            total_days = calculate_total_days(
                start_date,
                end_date
            )

            # =========================
            # AGREGAR DATA
            # =========================

            all_data.append({

                "Documento": identification,

                "Paciente": patient_name,

                "CONCATENAR":
                    f"{identification} {patient_name}"
                    if identification and patient_name
                    else None,

                "Fecha de Inicio": start_date,

                "Fecha de Fin": end_date,

                "Total de Días": total_days,

                "Dias": None,

                "EPS": eps,

                "Observacion": None,

                "DX": diagnosis,

                "FECHA": date_of_attention,

                "ORDEN": order,

                "PRORROGA": prorogation,

                "Tipo Incapacidad": tipo_incapacidad,

                "Grupo Servicio": grupo_servicio

            })

            pdf_document.close()

        # =========================
        # DATAFRAME
        # =========================

        df = pd.DataFrame(all_data)

        # =========================
        # HOMOLOGAR CIE10
        # =========================

        df = pd.merge(

            df,

            cie10_df,

            left_on='DX',

            right_on='Codigo',

            how='left'

        )

        # Concatenar DX + Nombre
        df['DX'] = (

            df['DX'].fillna('')

            + ' '

            + df['Nombre'].fillna('')

        )

        # Eliminar columnas
        df = df.drop(
            columns=['Nombre', 'Codigo'],
            errors='ignore'
        )

        # Fecha envío
        df.insert(

            0,

            'Fecha envio dra',

            datetime.now().strftime("%d/%m/%Y")

        )

        # =========================
        # USAR PLANTILLA
        # =========================

        # Cargar plantilla
        wb = load_workbook("PLANTILLA.xlsx")

        # Hoja INFO
        ws = wb["INFO"]

        # =========================
        # LIMPIAR DATOS ANTERIORES
        # =========================

        if ws.max_row > 1:

            ws.delete_rows(
                2,
                ws.max_row
            )

        # =========================
        # INSERTAR DATAFRAME
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
        # GUARDAR EN MEMORIA
        # =========================

        output = BytesIO()

        wb.save(output)

        output.seek(0)

        # =========================
        # RETORNAR EXCEL
        # =========================

        return send_file(

            output,

            download_name=f"Formato Rechazos Sura {datetime.now().strftime('%d-%m-%Y')}.xlsx",

            as_attachment=True,

            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        )

    except Exception as e:

        return str(e), 500

# =========================
# RUN
# =========================

if __name__ == "__main__":

    app.run(debug=True)
