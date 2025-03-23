from flask import Flask, request, render_template, send_file
import os
import pandas as pd
from fpdf import FPDF
import zipfile

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = 'pdfs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)


# Funzione che crea un PDF da una singola riga del DataFrame
def genera_pdf(row, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    # Supponendo che ci siano almeno 4 colonne
    for colonna in row.index[:4]:
        valore = row[colonna]
        pdf.cell(200, 10, txt=f"{colonna}: {valore}", ln=True)

    pdf.output(filename)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file_excel']

        if file and file.filename.endswith('.xlsx'):
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath)

            # Pulisci vecchi PDF
            for f in os.listdir(PDF_FOLDER):
                os.remove(os.path.join(PDF_FOLDER, f))

            # Genera PDF per ogni riga
            pdf_files = []
            for idx, row in df.iterrows():
                pdf_filename = os.path.join(PDF_FOLDER, f'riga_{idx+1}.pdf')
                genera_pdf(row, pdf_filename)
                pdf_files.append(pdf_filename)

            # Crea un archivio ZIP con tutti i PDF
            zip_path = os.path.join(PDF_FOLDER, 'pdf_generati.zip')
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for pdf in pdf_files:
                    zipf.write(pdf, os.path.basename(pdf))

            return send_file(zip_path, as_attachment=True)
        else:
            return "Formato file non valido. Usa un file .xlsx", 400

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)

# requirements.txt
# flask
# pandas
# openpyxl
# fpdf
