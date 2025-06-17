import os
import uuid
import time  # เพิ่มเข้ามาเพื่อใช้หน่วงเวลา
import fitz  # PyMuPDF
from docx import Document
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
MAX_FILE_SIZE = 30 * 1024 * 1024 # จำกัดไฟล์ 30MB

# --- App Initialization ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE
CORS(app)

# --- สร้างโฟลเดอร์ที่จำเป็น ---
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

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

    if conversion_type == 'pdf-to-word':
        original_filename = secure_filename(file.filename)
        task_id = str(uuid.uuid4())
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}.pdf")
        
        base_name = os.path.splitext(original_filename)[0]
        output_filename_user = f"{base_name}_sandbox.docx"
        output_filename_server = f"{task_id}.docx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename_server)

        try:
            file.save(input_path)
            
            pdf_document = fitz.open(input_path)
            word_document = Document()
            
            for page in pdf_document:
                text = page.get_text("text")
                word_document.add_paragraph(text)
                word_document.add_page_break()

            word_document.save(output_path)
            pdf_document.close()
            
            # --- หน่วงเวลา 3 วินาที เพื่อจำลองการทำงาน ---
            time.sleep(3)
            
            # ลบไฟล์ input ทันทีเพื่อประหยัดพื้นที่
            os.remove(input_path)

            download_url = f"/api/download/{output_filename_server}?filename={output_filename_user}"
            return jsonify({
                "message": "Conversion successful!",
                "download_url": download_url
            })

        except Exception as e:
            print(f"ERROR: {str(e)}")
            return jsonify({"error": "An internal error occurred."}), 500

    elif conversion_type == 'word-to-pdf':
         return jsonify({"error": "Word to PDF conversion is temporarily disabled."}), 501
    
    return jsonify({"error": "Invalid conversion type."}), 400

@app.route('/api/download/<server_filename>')
def download_file(server_filename):
    display_filename = request.args.get('filename', 'converted_file_sandbox.docx')
    try:
        return send_from_directory(
            app.config['UPLOAD_FOLDER'],
            server_filename,
            as_attachment=True,
            download_name=display_filename
        )
    finally:
        # พึ่งพาระบบไฟล์ของ Render ในการลบไฟล์ที่เหลือเพื่อลดความซับซ้อน
        pass

@app.route('/')
def index():
    return "File Conversion API (Fake Log Edition) is running!"

if __name__ == '__main__':
    app.run(debug=True, port=5000)
