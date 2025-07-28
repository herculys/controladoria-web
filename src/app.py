from flask import Flask, render_template, request, send_file
import os
from controladoria import process_files

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        panel_file = request.files['panel_file']
        office_file = request.files['office_file']
        intervalo_data = request.form['date_range']

        os.makedirs('uploads', exist_ok=True)

        panel_file_path = os.path.join('uploads', panel_file.filename)
        office_file_path = os.path.join('uploads', office_file.filename)

        panel_file.save(panel_file_path)
        office_file.save(office_file_path)

        relatorio_path = process_files(panel_file_path, office_file_path, intervalo_data)
        return send_file(relatorio_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)