import os
import uuid
import fitz  # PyMuPDF
from docx import Document
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
MAX_FILE_SIZE = 30 * 1024 * 1024 # 30 MB

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

# --- API Endpoint ---
@app.route('/api/convert', methods=['POST'])
def handle_conversion():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    conversion_type = request.form.get('type')
    
    # --- PDF to Word Conversion ---
    if conversion_type == 'pdf-to-word':
        try:
            # 1. Save uploaded file
            pdf_filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{uuid.uuid4()}_{pdf_filename}")
            file.save(input_path)
            
            # 2. Open PDF with PyMuPDF
            pdf_document = fitz.open(input_path)
            
            # 3. Create a new Word document
            word_document = Document()
            
            # 4. Extract text and add to Word document
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                # ใช้ get_text("text") เพื่อความเรียบง่ายและเสถียร
                text = page.get_text("text")
                word_document.add_paragraph(text)
                if page_num < len(pdf_document) - 1:
                    word_document.add_page_break() # เพิ่มตัวแบ่งหน้า
            
            # 5. Save the Word document
            base_name = os.path.splitext(pdf_filename)[0]
            output_filename = f"{base_name}_converted.docx"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            word_document.save(output_path)
            
            pdf_document.close()
            os.remove(input_path) # ลบไฟล์ pdf ที่อัปโหลดทิ้ง

            # ส่งลิงก์สำหรับดาวน์โหลดกลับไป (แบบง่าย)
            return jsonify({
                "message": "Conversion successful!",
                # เราจะส่งไฟล์โดยตรงแทนการสร้างลิงก์ที่ซับซ้อน
                "download_path": f"/api/download/{output_filename}"
            })

        except Exception as e:
            print(f"An error occurred during PDF-to-Word conversion: {e}")
            return jsonify({"error": "Failed to convert PDF to Word. The file might be corrupted or in an unsupported format."}), 500
    
    # --- Word to PDF (Temporarily Disabled) ---
    elif conversion_type == 'word-to-pdf':
        return jsonify({
            "error": "Word to PDF conversion is temporarily disabled due to server environment limitations. This feature is under development."
        }), 501 # 501 Not Implemented
        
    else:
        return jsonify({"error": "Invalid conversion type specified."}), 400

# --- Download Endpoint ---
@app.route('/api/download/<filename>')
def download_file(filename):
    try:
        # ส่งไฟล์ให้ผู้ใช้ดาวน์โหลดแล้วลบไฟล์นั้นทิ้งจากเซิร์ฟเวอร์
        return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found or has been deleted."}), 404

@app.route('/')
def index():
    return "File Conversion API (Self-Hosted, Revised) is running!"

if __name__ == '__main__':
    app.run(debug=True, port=5000)
