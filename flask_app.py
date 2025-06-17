import os
import uuid
import time
import threading
from flask import Flask, request, send_from_directory, jsonify, Response
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
from pdf2docx import Converter as PDF_Converter
from docx2pdf import convert as WORD_Converter

# --- Configuration ---
# 1. โฟลเดอร์สำหรับเก็บไฟล์
UPLOAD_FOLDER = 'uploads'
LOG_FOLDER = 'logs'
# 2. จำกัดขนาดไฟล์ 30MB (30 * 1024 * 1024 bytes)
MAX_FILE_SIZE = 30 * 1024 * 1024
# 3. นามสกุลไฟล์ที่อนุญาต
ALLOWED_EXTENSIONS = {'pdf', 'docx'}
# 4. เวลาที่จะเก็บไฟล์ไว้ก่อนลบ (วินาที)
CLEANUP_DELAY_SECONDS = 300 # 5 นาที

# --- App Initialization ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['LOG_FOLDER'] = LOG_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE
CORS(app) # อนุญาตการเรียก API จากโดเมนอื่น (WordPress ของเรา)

# --- สร้างโฟลเดอร์ที่จำเป็น ---
for folder in [UPLOAD_FOLDER, LOG_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# --- Helper Classes & Functions ---

class LogStreamer:
    """คลาสสำหรับช่วยเขียน Log ลงไฟล์ และอ่านเพื่อส่งให้ Client"""
    def __init__(self, task_id):
        self.log_file_path = os.path.join(app.config['LOG_FOLDER'], f"{task_id}.log")
        # สร้างไฟล์ log ว่างๆ ขึ้นมาก่อน
        with open(self.log_file_path, 'w', encoding='utf-8') as f:
            f.write('')

    def log(self, message):
        """เขียนข้อความลงในไฟล์ log พร้อม timestamp"""
        timestamp = time.strftime('%H:%M:%S')
        with open(self.log_file_path, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")
        print(f"Logged for {self.log_file_path}: {message}")

    def stream(self):
        """Generator function ที่จะคอยอ่าน log ใหม่ๆ จากไฟล์ส่งออกไป"""
        last_pos = 0
        while True:
            with open(self.log_file_path, 'r', encoding='utf-8') as f:
                f.seek(last_pos)
                new_lines = f.readlines()
                last_pos = f.tell()

                for line in new_lines:
                    line = line.strip()
                    if line:
                        # ส่งข้อมูลในรูปแบบ Server-Sent Event (SSE)
                        yield f"data: {line}\n\n"
                
                # ตรวจหาข้อความสุดท้ายเพื่อหยุดการ stream
                if "PROCESS_COMPLETE" in " ".join(new_lines):
                    yield f"data: [DONE]\n\n"
                    break
            time.sleep(0.5) # หน่วงเวลาเล็กน้อยเพื่อลดภาระ CPU


def cleanup_files(paths_to_delete):
    """ฟังก์ชันสำหรับลบไฟล์ทิ้งใน background thread เพื่อไม่ให้ client ต้องรอ"""
    def task():
        print(f"Cleanup scheduled in {CLEANUP_DELAY_SECONDS} seconds for: {paths_to_delete}")
        time.sleep(CLEANUP_DELAY_SECONDS)
        for path in paths_to_delete:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    print(f"Successfully cleaned up: {path}")
            except Exception as e:
                print(f"Error during cleanup of {path}: {e}")

    thread = threading.Thread(target=task)
    thread.start()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- Error Handlers ---
@app.errorhandler(413)
@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    return jsonify({"error": f"File is too large. Maximum size is {MAX_FILE_SIZE / 1024 / 1024:.0f} MB."}), 413

# --- API Endpoints ---

@app.route('/api/convert', methods=['POST'])
def handle_conversion():
    # 1. ตรวจสอบไฟล์
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    
    conversion_type = request.form.get('type') # 'pdf-to-word' or 'word-to-pdf'
    if not conversion_type:
        return jsonify({"error": "Conversion type not specified"}), 400

    if file and allowed_file(file.filename):
        # 2. สร้าง ID ที่ไม่ซ้ำกันสำหรับงานนี้
        task_id = str(uuid.uuid4())
        logger = LogStreamer(task_id)

        try:
            # 3. เตรียมชื่อไฟล์และ path
            original_filename = secure_filename(file.filename)
            base_name, extension = os.path.splitext(original_filename)
            
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}{extension}")
            file.save(input_path)
            logger.log(f"File '{original_filename}' uploaded successfully.")

            # 4. กำหนดชื่อไฟล์ output
            output_extension = '.docx' if conversion_type == 'pdf-to-word' else '.pdf'
            output_filename = f"{base_name}_sandbox{output_extension}"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_output{output_extension}")

            # 5. เริ่มกระบวนการแปลงไฟล์
            logger.log(f"Starting conversion: {conversion_type.replace('-', ' ')}...")

            if conversion_type == 'pdf-to-word':
                if extension.lower() != '.pdf':
                    raise ValueError("Invalid file type for PDF to Word conversion.")
                cv = PDF_Converter(input_path)
                cv.convert(output_path, start=0, end=None)
                cv.close()
            elif conversion_type == 'word-to-pdf':
                if extension.lower() != '.docx':
                    raise ValueError("Invalid file type for Word to PDF conversion.")
                WORD_Converter(input_path, output_path)
            else:
                raise ValueError("Invalid conversion type.")

            logger.log("Conversion completed successfully!")
            logger.log(f"Output file: {output_filename}")
            logger.log("PROCESS_COMPLETE") # สัญญาณบอกว่าจบงานแล้ว

            # 6. ส่ง URL สำหรับดาวน์โหลดและ stream log กลับไป
            # เราจะใช้ชื่อไฟล์ output ที่เราสร้าง (เช่น mydoc_sandbox.docx) ตอนดาวน์โหลด
            # แต่ตัวไฟล์บน server จะเป็น task_id เพื่อให้จัดการง่าย
            download_url = f"/api/download/{task_id}_output{output_extension}?filename={output_filename}"
            
            return jsonify({
                "message": "Processing started.",
                "log_stream_url": f"/api/stream-logs/{task_id}",
                "download_url": download_url
            })

        except Exception as e:
            logger.log(f"ERROR: {str(e)}")
            logger.log("PROCESS_COMPLETE")
            print(f"An error occurred: {e}")
            return jsonify({"error": f"An error occurred: {str(e)}"}), 500
        finally:
            # กำหนดการลบไฟล์ input/output/log ในภายหลัง
            cleanup_files([
                input_path, 
                output_path,
                os.path.join(app.config['LOG_FOLDER'], f"{task_id}.log")
            ])
    
    return jsonify({"error": "Invalid file type."}), 400

@app.route('/api/stream-logs/<task_id>')
def stream_logs(task_id):
    """Endpoint สำหรับส่ง log แบบ real-time ด้วย SSE"""
    logger = LogStreamer(task_id)
    # ใช้ Response object ของ Flask เพื่อ stream ข้อมูล
    return Response(logger.stream(), mimetype='text/event-stream')

@app.route('/api/download/<server_filename>')
def download_file(server_filename):
    """Endpoint สำหรับดาวน์โหลดไฟล์ที่แปลงเสร็จแล้ว"""
    # รับชื่อไฟล์ที่ต้องการให้ user เห็นจาก query string
    display_filename = request.args.get('filename', 'converted_file_sandbox')
    return send_from_directory(
        app.config['UPLOAD_FOLDER'],
        server_filename,
        as_attachment=True,
        download_name=display_filename # ตั้งชื่อไฟล์ตอนดาวน์โหลด
    )

@app.route('/')
def index():
    return "Advanced File Conversion API v2 is running!"

# ไม่จำเป็นสำหรับ PythonAnywhere แต่มีไว้ทดสอบบนเครื่อง
if __name__ == '__main__':
    app.run(debug=True, port=5000, threaded=True)
