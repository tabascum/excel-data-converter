from flask import Flask, render_template, request, redirect, url_for, send_file
import data_processor
import os

app = Flask(__name__)

# Define the download directory
DOWNLOAD_DIRECTORY = 'downloads'

# Ensure the download directory exists
if not os.path.exists(DOWNLOAD_DIRECTORY):
    os.makedirs(DOWNLOAD_DIRECTORY)

@app.route("/")
def homepage():
    return render_template("index.html")

@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return "No file selected"

        file = request.files['file']

        if file.filename == '':
            return "No file selected"

        if file:
            # Save the uploaded file to the downloads directory
            file_path = os.path.join(DOWNLOAD_DIRECTORY, file.filename)
            file.save(file_path)

            # Process the uploaded file with the download_directory argument
            processed_file_path = data_processor.process_excel(file_path, DOWNLOAD_DIRECTORY)

            # Construct download URL
            download_url = url_for('download', filename=os.path.basename(processed_file_path))

            return render_template('processed.html', download_url=download_url)
    except Exception as e:
        return str(e)


@app.route('/download/<filename>')
def download(filename):
    return send_file(os.path.join(DOWNLOAD_DIRECTORY, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(host='127.0.0.1',port=8000,debug=True)

