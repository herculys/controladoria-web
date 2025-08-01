
import sys
import subprocess

def install_requirements():
    try:
        import flask
        import pandas
        import openpyxl
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

install_requirements()

from flask import Flask, render_template, request, send_file
from controladoria import process_files

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        panel_file = request.files['panel_file']
        office_file = request.files['office_file']

        # Passe os objetos FileStorage diretamente
        relatorio_bytesio = process_files(panel_file, office_file)
        return send_file(
            relatorio_bytesio,
            as_attachment=True,
            download_name="RECOLHIMENTO CONTROLADORIA.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)