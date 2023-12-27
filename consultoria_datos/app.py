from flask import Flask, render_template, request
import pandas as pd
import os
import uuid
import gspread
from google.oauth2 import service_account
from gspread_dataframe import set_with_dataframe

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

# Configuración de Google Sheets
CREDENTIALS_FILE = 'C:\\Users\\ruben\\Desktop\\PRUEBA_BD\\base-de-datos-409414-b1b5369a48bf.json'
SPREADSHEET_ID = '1bSL2Z05mmprzZQRWTrNJrHKzPjIkVB-9QHKj9qirrws'  # ID de tu hoja de cálculo

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_unique_filename():
    return f'resultado_{uuid.uuid4().hex}.xlsx'

def write_to_google_sheets(df):
    credentials = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            # Otros alcances si son necesarios
        ]
    )
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)
    worksheet = spreadsheet.get_worksheet(0)  # Puedes cambiar el índice según la hoja que quieras usar

    # Obtener la última fila ocupada en la hoja de cálculo
    last_row = len(worksheet.get_all_values()) + 1

    # Añadir datos a la hoja de cálculo sin truncar y sin incluir encabezados ni índices
    df_no_headers = df.iloc[1:] if not df.empty and df.iloc[0].notna().all() else df

    # Convertir columnas que parecen ser fechas a formato adecuado
    date_columns = df_no_headers.select_dtypes(include=['datetime']).columns
    for col in date_columns:
        df_no_headers[col] = df_no_headers[col].dt.strftime('%Y-%m-%d %H:%M:%S')

    set_with_dataframe(worksheet, df_no_headers, row=last_row, include_index=False, resize=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'archivo' not in request.files:
        return render_template('index.html', error="No se seleccionó ningún archivo.")

    archivos = request.files.getlist('archivo')

    hoja_index = int(request.form['hoja'])  # Obtén el índice de la hoja seleccionada

    for archivo in archivos:
        if archivo.filename == '':
            return render_template('index.html', error="Nombre de archivo vacío.")

        if not allowed_file(archivo.filename):
            return render_template('index.html', error="Tipo de archivo no permitido. Solo se permiten archivos Excel (xlsx, xls) o CSV.")

        try:
            # Leer el archivo Excel y seleccionar la hoja específica
            df = pd.read_excel(archivo, sheet_name=hoja_index, header=None, engine='openpyxl')

            # Realizar operaciones o consultas con los datos (opcional)

            # Guardar en Google Sheets
            write_to_google_sheets(df)

        except pd.errors.EmptyDataError:
            return render_template('index.html', error="El archivo está vacío.")

        except Exception as e:
            return render_template('index.html', error=f"Error al procesar el archivo: {str(e)}")

    return render_template('index.html', success="Datos cargados exitosamente en Google Sheets.")

if __name__ == '__main__':
    app.run(debug=True)
