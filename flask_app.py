import os
import uuid
import time
import threading
import fitz  # PyMuPDF
from docx import Document
from flask import Flask, request, jsonify, Response, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
LOG_FOLDER = 'logs'
MAX_FILE_SIZE = 30 * 1024 * 1024
CLEANUP_DELAY_SECONDS = 300 # 5 นาที

# --- App Initialization ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['LOG_FOLDER'] = LOG_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE
CORS(app)

# --- สร้างโฟลเดอร์ที่จำเป็น ---
for folder in [UPLOAD_FOLDER, LOG_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# --- Helper Classes & Functions ---
class LogStreamer:
    """คลาสสำหรับช่วยเขียน Log ลงไฟล์ และอ่านเพื่อส่งให้ Client"""
    def __init__(self, task_id):
        self.log_file_path = os.path.join(app.config['LOG_FOLDER'], f"{task_id}.log")
        with open(self.log_file_path, 'w', encoding='utf-8') as f:
            f.write('')

    def log(self, message):
        timestamp = time.strftime('%H:%M:%S')
        with open(self.log_file_path, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")
        print(f"Logged for {self.log_file_path}: {message}")

    def stream(self):
        last_pos = 0
        while True:
            with open(self.log_file_path, 'r', encoding='utf-8') as f:
                f.seek(last_pos)
                new_lines = f.readlines()
                last_pos = f.tell()
                for line in new_lines:
                    line = line.strip()
                    if line:
                        yield f"data: {line}\n\n"
                if "PROCESS_COMPLETE" in " ".join(new_lines):
                    yield f"data: [DONE]\n\n"
                    break
            time.sleep(0.5)

def cleanup_files(paths_to_delete):
    def task():
        time.sleep(CLEANUP_DELAY_SECONDS)
        for path in paths_to_delete:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                print(f"Error during cleanup of {path}: {e}")
    thread = threading.Thread(target=task)
    thread.start()

# --- Error Handlers ---
@app.errorhandler(413)
@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    return jsonify({"error": f"File is too large. Maximum size is {MAX_FILE_SIZE / 1024 / 1024:.0f} MB."}), 413

# --- API Endpoints ---
@app.route('/api/convert', methods=['POST'])
def handle_conversion():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    conversion_type = request.form.get('type')
    task_id = str(uuid.uuid4())
    logger = LogStreamer(task_id)

    if conversion_type == 'pdf-to-word':
        original_filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}.pdf")
        
        # ใช้ชื่อไฟล์ output ที่เราสร้าง (_sandbox)
        base_name = os.path.splitext(original_filename)[0]
        output_filename_user = f"{base_name}_sandbox.docx"
        output_filename_server = f"{task_id}.docx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename_server)

        try:
            file.save(input_path)
            logger.log(f"File '{original_filename}' received. Task ID: {task_id}")
            logger.log("Initializing conversion engine (PyMuPDF)...")
            
            pdf_document = fitz.open(input_path)
            word_document = Document()
            
            logger.log(f"PDF contains {len(pdf_document)} pages. Starting text extraction...")
            
            for i, page in enumerate(pdf_document):
                text = page.get_text("text")
                word_document.add_paragraph(text)
                word_document.add_page_break()
                logger.log(f"-> Extracted text from page {i+1}/{len(pdf_document)}")

            logger.log("Assembling Word document...")
            word_document.save(output_path)
            pdf_document.close()
            
            logger.log("Conversion complete!")
            logger.log(f"Output file ready: {output_filename_user}")
            logger.log("PROCESS_COMPLETE")

            download_url = f"/api/download/{output_filename_server}?filename={output_filename_user}"
            return jsonify({
                "message": "Processing started.",
                "log_stream_url": f"/api/stream-logs/{task_id}",
                "download_url": download_url
            })

        except Exception as e:
            error_message = f"ERROR: {str(e)}"
            logger.log(error_message)
            logger.log("PROCESS_COMPLETE")
            print(error_message)
            return jsonify({"error": "An internal error occurred."}), 500
        finally:
             cleanup_files([input_path, output_path, os.path.join(app.config['LOG_FOLDER'], f"{task_id}.log")])

    elif conversion_type == 'word-to-pdf':
         return jsonify({"error": "Word to PDF conversion is temporarily disabled."}), 501
    
    return jsonify({"error": "Invalid conversion type."}), 400

@app.route('/api/stream-logs/<task_id>')
def stream_logs(task_id):
    logger = LogStreamer(task_id)
    return Response(logger.stream(), mimetype='text/event-stream')

@app.route('/api/download/<server_filename>')
def download_file(server_filename):
    display_filename = request.args.get('filename', 'converted_file_sandbox.docx')
    return send_from_directory(
        app.config['UPLOAD_FOLDER'],
        server_filename,
        as_attachment=True,
        download_name=display_filename
    )

@app.route('/')
def index():
    return "File Conversion API (Self-Hosted, SSE) is running!"

if __name__ == '__main__':
    app.run(debug=True, port=5000, threaded=True)
