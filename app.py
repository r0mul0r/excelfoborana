import os
import re
import io
import pandas as pd
from flask import Flask, request, send_file, render_template, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB max
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def parse_monto(value):
    """
    Convierte distintos formatos de monto a float.
    Soporta: '10.000,00' | '10,000.00' | '10000' | '10000.00' | 1000 (número)
    """
    if value is None:
        return 0.0

    # Si ya es numérico, devolver directo
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    # Quitar símbolo de moneda y espacios
    text = re.sub(r'[^\d.,\-]', '', text)

    if not text or text in ('.', ',', '-'):
        return 0.0

    # Detectar formato: si tiene punto como separador de miles y coma decimal → '10.000,00'
    if re.search(r'\d\.\d{3}', text) and ',' in text:
        # Formato europeo/español: 10.000,00
        text = text.replace('.', '').replace(',', '.')
    elif ',' in text and '.' in text:
        # Ambos presentes: determinar cuál va último
        last_dot = text.rfind('.')
        last_comma = text.rfind(',')
        if last_comma > last_dot:
            # Coma es decimal: 10.000,50
            text = text.replace('.', '').replace(',', '.')
        else:
            # Punto es decimal: 10,000.50
            text = text.replace(',', '')
    elif ',' in text:
        # Solo coma: puede ser decimal o miles
        comma_parts = text.split(',')
        if len(comma_parts) == 2 and len(comma_parts[1]) == 3:
            # Coma como miles: 10,000
            text = text.replace(',', '')
        else:
            # Coma como decimal: 10,5
            text = text.replace(',', '.')

    try:
        return float(text)
    except ValueError:
        return 0.0


def find_columns(df):
    """
    Busca las columnas USUARIO y MONTO RECARGADO de forma flexible
    (sin importar mayúsculas, espacios extra, acentos, etc.)
    """
    col_map = {}
    for col in df.columns:
        normalized = str(col).strip().upper()
        if normalized in ('USUARIO', 'USER', 'USUARIOS'):
            col_map['usuario'] = col
        elif any(k in normalized for k in ('MONTO', 'IMPORTE', 'AMOUNT')):
            col_map['monto'] = col

    return col_map.get('usuario'), col_map.get('monto')


def process_excel(file_bytes):
    """
    Lee el Excel, agrupa por usuario y calcula totales.
    Retorna bytes del Excel resultante.
    """
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    except Exception:
        df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')

    col_usuario, col_monto = find_columns(df)

    if col_usuario is None:
        raise ValueError("No se encontró la columna USUARIO en el archivo.")
    if col_monto is None:
        raise ValueError("No se encontró la columna MONTO RECARGADO en el archivo.")

    df['_usuario'] = df[col_usuario].astype(str).str.strip()
    df['_monto'] = df[col_monto].apply(parse_monto)

    # Filtrar filas vacías o nulas
    df = df[df['_usuario'].notna() & (df['_usuario'] != '') & (df['_usuario'] != 'nan')]

    resultado = df.groupby('_usuario').agg(
        CANTIDAD_RECARGAS=('_monto', 'count'),
        MONTO_TOTAL=('_monto', 'sum')
    ).reset_index()

    resultado.columns = ['USUARIO', 'CANTIDAD DE RECARGAS', 'MONTO TOTAL RECARGADO']
    resultado = resultado.sort_values('MONTO TOTAL RECARGADO', ascending=False)

    return build_output_excel(resultado)


def build_output_excel(df):
    """Genera el Excel de salida con formato."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Recargas"

    # Colores
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    alt_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    total_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    headers = ['USUARIO', 'CANTIDAD DE RECARGAS', 'MONTO TOTAL RECARGADO']
    col_widths = [25, 22, 25]

    # Encabezados
    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 30

    # Datos
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        fill = alt_fill if row_idx % 2 == 0 else None

        cell_usuario = ws.cell(row=row_idx, column=1, value=row[0])
        cell_usuario.border = thin_border
        cell_usuario.alignment = Alignment(horizontal='left', vertical='center')
        if fill:
            cell_usuario.fill = fill

        cell_cant = ws.cell(row=row_idx, column=2, value=int(row[1]))
        cell_cant.border = thin_border
        cell_cant.alignment = Alignment(horizontal='center', vertical='center')
        if fill:
            cell_cant.fill = fill

        cell_monto = ws.cell(row=row_idx, column=3, value=row[2])
        cell_monto.border = thin_border
        cell_monto.alignment = Alignment(horizontal='right', vertical='center')
        cell_monto.number_format = '#,##0.00'
        if fill:
            cell_monto.fill = fill

    # Fila totales
    total_row = len(df) + 2
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=total_row, column=1).fill = total_fill
    ws.cell(row=total_row, column=1).border = thin_border
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=total_row, column=2, value=int(df['CANTIDAD DE RECARGAS'].sum())).font = Font(bold=True, color="FFFFFF")
    ws.cell(row=total_row, column=2).fill = total_fill
    ws.cell(row=total_row, column=2).border = thin_border
    ws.cell(row=total_row, column=2).alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=total_row, column=3, value=df['MONTO TOTAL RECARGADO'].sum()).font = Font(bold=True, color="FFFFFF")
    ws.cell(row=total_row, column=3).fill = total_fill
    ws.cell(row=total_row, column=3).border = thin_border
    ws.cell(row=total_row, column=3).number_format = '#,##0.00'
    ws.cell(row=total_row, column=3).alignment = Alignment(horizontal='right', vertical='center')

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar():
    if 'archivo' not in request.files:
        return jsonify({'error': 'No se envió ningún archivo.'}), 400

    file = request.files['archivo']

    if file.filename == '':
        return jsonify({'error': 'No se seleccionó ningún archivo.'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Formato no soportado. Solo se aceptan archivos .xlsx o .xls'}), 400

    try:
        file_bytes = file.read()
        result_bytes = process_excel(file_bytes)

        output = io.BytesIO(result_bytes)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='resumen_recargas.xlsx'
        )
    except ValueError as e:
        return jsonify({'error': str(e)}), 422
    except Exception as e:
        return jsonify({'error': f'Error procesando el archivo: {str(e)}'}), 500


if __name__ == '__main__':
    app.run(debug=True)
