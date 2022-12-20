import os
from convert import *

from flask import Flask, render_template, request, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
from flask import send_from_directory

UPLOAD_FOLDER = './files/uploads'
DOWNLOAD_FOLDER = './files/downloads/'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.secret_key = 'v3islvno;i@#5xzcv'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/download')
def download(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'],
                           filename, as_attachment=True)


@app.route('/', methods=['GET', 'POST'])
def converter():
    download = False
    if request.method == 'POST':
        if 'excelfile' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['excelfile']


        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            print(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename)))
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename)))
            # Fetch the input file

            wb = fetch_input_workbook(secure_filename(file.filename))

            # Pull all the existing sheet names
            sheets = wb.sheetnames

            # Process the latest sheet (Pass in sheets and function will select latest)
            output_file = process_sheets(wb, sheets)
            #session['output_file_name'] = DOWNLOAD_FOLDER + secure_filename(output_file)

            f = os.path.join(DOWNLOAD_FOLDER, output_file)

            return send_from_directory(app.config['DOWNLOAD_FOLDER'],
                                output_file, as_attachment=True)
    else :

        return render_template('converter.html')



if __name__ == '__main__':
    app.run()
