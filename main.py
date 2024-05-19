import os
import io
from flask import Flask, flash, request, redirect, send_file
from easyocr import Reader
from utils import create_excel_file, process_image

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

app = Flask(__name__)

reader = Reader(['en'])

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        results = []
        for file in request.files.getlist('file'):
            if file and allowed_file(file.filename):                
                results.append(process_image(reader, file, 1550, 300, 1910, 380))
        
        excel_bytes_array = create_excel_file(results)

        return send_file(excel_bytes_array,
                        mimetype='application/vnd.ms-excel',
                        download_name='output.xlsx')
    return '''
    <!doctype html>
    <title>Prize Bonds OCR</title>
    <h1>Upload Scanned Images</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file multiple=true>
      <input type=submit value=Upload>
    </form>
    '''