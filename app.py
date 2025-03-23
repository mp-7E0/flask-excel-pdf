from flask import Flask, request, render_template, send_file
import os
import pandas as pd
from fpdf import FPDF
import zipfile
from azure.storage.blob import BlobServiceClient

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Azure Blob Storage configuration
AZURE_STORAGE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=storagefrontend;AccountKey=QcqB/SGawtZU5hKgZOOFvYSvvjIdGcYbymA8XRff0hAF07+h8QgtrXCJXd+n67E53WE/S4LG+ESF+AStxOKCtA==;EndpointSuffix=core.windows.net"
CONTAINER_NAME = "flask-excel-pdf"
blob_service_client = BlobServiceClient.from_connection_string(AZURE_STORAGE_CONNECTION_STRING)

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

# Funzione per caricare il file PDF su Azure Blob Storage
def upload_to_blob(local_file_path, blob_name):
    blob_client = blob_service_client.get_blob_client(container=CONTAINER_NAME, blob=blob_name)
    with open(local_file_path, "rb") as data:
        blob_client.upload_blob(data, overwrite=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file_excel']

        if file and file.filename.endswith('.xlsx'):
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath)

            # Genera PDF per ogni riga e carica su Azure Blob Storage
            pdf_files = []
            for idx, row in df.iterrows():
                local_pdf_filename = f'riga_{idx+1}.pdf'
                local_pdf_filepath = os.path.join(UPLOAD_FOLDER, local_pdf_filename)
                genera_pdf(row, local_pdf_filepath)

                # Carica PDF su Azure Blob Storage
                upload_to_blob(local_pdf_filepath, local_pdf_filename)
                pdf_files.append(local_pdf_filename)

            return "PDF generati e caricati su Azure Blob Storage con successo!"
        else:
            return "Formato file non valido. Usa un file .xlsx", 400

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
