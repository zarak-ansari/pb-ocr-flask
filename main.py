import os
import io
from flask import Flask, flash, request, redirect, send_file, render_template
from easyocr import Reader
from utils import create_excel_file, process_image

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024
reader = Reader(['en'])

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        left = int(request.form.get('left'))
        top = int(request.form.get('top'))
        right = int(request.form.get('right'))
        bottom = int(request.form.get('bottom'))
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        results = []
        for file in request.files.getlist('file'):
            if file and allowed_file(file.filename):                
                results.append(process_image(reader, file, left, top, right, bottom))
        
        excel_bytes_array = create_excel_file(results)

        return send_file(excel_bytes_array,
                        mimetype='application/vnd.ms-excel',
                        download_name='output.xlsx')
    return render_template('index.html')