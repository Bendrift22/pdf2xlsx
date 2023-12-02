from flask import Flask, render_template, request, send_file
import os
from main import main

app = Flask(__name__)

UPLOAD_FOLDER = 'upload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        print("OK")
        return render_template('index.html', message='Aucun fichier trouvé.')

    file = request.files['file']

    if file.filename == '':
        return render_template('index.html', message='Aucun fichier sélectionné.')

    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'uploaded_file.pdf')
        file.save(file_path)
        main()
        return render_template('index.html', message='Fichier téléchargé avec succès.', file_path=file_path)

@app.route('/download')
def download_file():
    file_path = "pdf.xlsx"
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
