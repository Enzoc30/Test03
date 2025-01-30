from flask import Flask, request, render_template, send_file
import pandas as pd
import re
import zipfile
import tempfile
import os
from PyPDF2 import PdfReader
from flask_cors import CORS 


app = Flask(__name__)

CORS(app)


def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def process_text_to_dataframe(text):
    lines = text.splitlines()
    data = []

    saldo_pattern = re.compile(r'^\s*SALDO ANTERIOR\s+([\d,]+\.\d{2})\s*$', re.IGNORECASE)
    transaction_pattern = re.compile(r'^(\d{2}[A-Z]{3})\s+(\d{2}[A-Z]{3})\s+(.*?)(\s+)([\d,]+\.\d{2})\s*$')

    for line in lines:
        line = line.strip()
        
        if saldo_match := saldo_pattern.match(line):
            amount = saldo_match.group(1).replace(',', '')
            data.append({'FECHA PROC.': None, 'FECHA VALOR': None, 'DESCRIPCION': 'SALDO ANTERIOR', 'CARGOS / DEBE': None, 'ABONOS / HABER': float(amount)})
            continue

        if trans_match := transaction_pattern.match(line):
            fecha_proc, fecha_valor, descripcion, espacios, monto = trans_match.groups()
            monto = float(monto.replace(',', ''))
            espacios_len = len(espacios)
            
            cargos = monto if espacios_len <= 30 else None
            abonos = monto if espacios_len > 30 else None
            descripcion = re.sub(r'\s+', ' ', descripcion).strip()
            
            data.append({'FECHA PROC.': fecha_proc, 'FECHA VALOR': fecha_valor, 'DESCRIPCION': descripcion, 'CARGOS / DEBE': cargos, 'ABONOS / HABER': abonos})

    return pd.DataFrame(data, columns=['FECHA PROC.', 'FECHA VALOR', 'DESCRIPCION', 'CARGOS / DEBE', 'ABONOS / HABER'])

def process_zip(zip_path):
    with tempfile.TemporaryDirectory() as temp_dir:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        all_data = []
        pdf_files = []

        for root, _, files in os.walk(temp_dir):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))

        pdf_files.sort(key=lambda x: os.path.basename(x).lower())

        for pdf_path in pdf_files:
            text = extract_text_from_pdf(pdf_path)
            df = process_text_to_dataframe(text)
            all_data.append(df)
            all_data.append(pd.DataFrame([['']*5], columns=df.columns))
            all_data.append(pd.DataFrame([['']*5], columns=df.columns))

        return pd.concat(all_data[:-2], ignore_index=True) if all_data else pd.DataFrame()

def save_to_excel(df, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
        workbook = writer.book
        worksheet = writer.sheets['Consolidado']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file uploaded', 400
        
        file = request.files['file']
        if file.filename == '':
            return 'No selected file', 400
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.zip') as temp_zip:
            file.save(temp_zip.name)
            temp_zip_path = temp_zip.name
        
        df_consolidado = process_zip(temp_zip_path)
        output_excel_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
        save_to_excel(df_consolidado, output_excel_path)
        os.remove(temp_zip_path)

        return send_file(output_excel_path, as_attachment=True, download_name='output.xlsx')
    
    return '''
    <!doctype html>
    <html>
        <body>
            <h1>Subir archivo ZIP</h1>
            <form method="post" enctype="multipart/form-data">
                <input type="file" name="file">
                <input type="submit" value="Procesar">
            </form>
        </body>
    </html>
    '''

if __name__ == '__main__':
    app.run(debug=True)
