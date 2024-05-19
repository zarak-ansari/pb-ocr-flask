import os
from flask import Flask, flash, request, redirect, url_for
from werkzeug.utils import secure_filename
from easyocr import Reader
from utils.ocr import process_image

UPLOAD_FOLDER = "C:\\Users\\zarak\\Workspace\\pb_ocr_flask\\uploads"
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}

app = Flask(__name__)

reader = Reader(['en'])


app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

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
        
        for file in request.files.getlist('file'):
            # If the user does not select a file, the browser submits an
            # empty file without a filename.
            if file and allowed_file(file.filename):
                cropped_image = process_image(reader, file, 1550, 300, 1910, 380)
                filename = secure_filename(file.filename)
                cropped_image.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        return redirect(url_for('files_uploaded'))
    return '''
    <!doctype html>
    <title>Upload new File</title>
    <h1>Upload new File</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file multiple=true>
      <input type=submit value=Upload>
    </form>
    '''

@app.route('/done')
def files_uploaded():
    return '''
    <!doctype html>
    <head>
    <title>Success!</title>
    </head>
    <body>
    <h1>Files uploaded</h1>
    </body>
    '''