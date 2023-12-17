import os
from pathlib import Path
from utils import search_files
from main import description_compare
from flask import Flask, render_template, request, send_file, send_from_directory

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

UPLOAD_FOLDER = Path(__file__).parent
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = ['csv']


@app.route("/")
def index():
    files = search_files(UPLOAD_FOLDER)
    return render_template("index.html", files=files)


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],
                               filename)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = file.filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            description_compare(file.filename)
            files = search_files(UPLOAD_FOLDER)
            return render_template("index.html", files=files)


@app.route('/download/<filename>')
def download(filename):
    return send_file(filename)


@app.route('/remove/<filename>')
def remove(filename):
    if os.path.exists(filename):
        os.remove(filename)
    files = search_files(UPLOAD_FOLDER)
    return render_template("index.html", files=files)


if __name__ == "__main__":
    app.run(debug=True)
