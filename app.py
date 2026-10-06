# -*- coding: utf-8 -*-
# FLASK_APP = app.py
# PDF TO WORD CONVERTER - WEB APP
# YÊU CẦU: pip install flask pdf2docx
# CHẠY: python app.py

import os
import uuid
from flask import Flask, request, render_template_string, send_file, jsonify, redirect, url_for
from werkzeug.utils import secure_filename
from pdf2docx import Converter

# ========== CẤU HÌNH ==========
app = Flask(__name__)
app.secret_key = 'PDF2WORD_SECRET_2026'

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# ========== HÀM KIỂM TRA ĐỊNH DẠNG ==========
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ========== HTML TEMPLATE ==========
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Chuyển đổi PDF sang Word Online</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: 'Segoe UI', sans-serif;
            padding: 20px;
        }
        .container-box {
            max-width: 700px;
            width: 100%;
            background: #ffffff;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
        }
        h1 {
            color: #333;
            font-weight: 700;
            text-align: center;
            margin-bottom: 10px;
        }
        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 30px;
        }
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 40px 20px;
            text-align: center;
            background: #f8f9ff;
            transition: all 0.3s;
            cursor: pointer;
            margin-bottom: 20px;
        }
        .upload-area:hover {
            border-color: #764ba2;
            background: #eef1ff;
        }
        .upload-area.dragover {
            border-color: #764ba2;
            background: #e0e5ff;
            transform: scale(1.02);
        }
        .upload-icon {
            font-size: 60px;
            color: #667eea;
            margin-bottom: 15px;
        }
        .file-info {
            display: none;
            background: #e8f5e9;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 20px;
            border-left: 4px solid #4caf50;
        }
        .file-info.show {
            display: block;
        }
        .btn-convert {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            color: white;
            padding: 15px 40px;
            border-radius: 30px;
            font-size: 16px;
            font-weight: 600;
            width: 100%;
            transition: transform 0.2s;
        }
        .btn-convert:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }
        .btn-convert:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }
        .loading.show {
            display: block;
        }
        .spinner-border {
            color: #667eea;
        }
        .result-card {
            display: none;
            background: #e8f5e9;
            border-radius: 10px;
            padding: 20px;
            margin-top: 20px;
            border-left: 4px solid #4caf50;
        }
        .result-card.show {
            display: block;
        }
        .error-card {
            display: none;
            background: #ffebee;
            border-radius: 10px;
            padding: 20px;
            margin-top: 20px;
            border-left: 4px solid #f44336;
        }
        .error-card.show {
            display: block;
        }
        .btn-download {
            background: #4caf50;
            border: none;
            color: white;
            padding: 12px 30px;
            border-radius: 25px;
            font-weight: 600;
            text-decoration: none;
            display: inline-block;
            margin-top: 10px;
        }
        .btn-download:hover {
            background: #45a049;
            color: white;
        }
        .features {
            display: flex;
            justify-content: space-around;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
        }
        .feature {
            text-align: center;
            font-size: 13px;
            color: #666;
        }
        .feature-icon {
            font-size: 24px;
            display: block;
            margin-bottom: 5px;
        }
    </style>
</head>
<body>
<div class="container-box">
    <h1>📄 → 📝 PDF sang Word</h1>
    <p class="subtitle">Chuyển đổi file PDF sang Word (.docx) miễn phí, nhanh chóng</p>

    <form id="uploadForm" enctype="multipart/form-data">
        <div class="upload-area" id="uploadArea">
            <div class="upload-icon">📁</div>
            <p><strong>Kéo thả file PDF vào đây</strong></p>
            <p style="color: #999; font-size: 14px;">hoặc click để chọn file</p>
            <p style="color: #999; font-size: 12px; margin-top: 10px;">Kích thước tối đa: 50MB</p>
            <input type="file" id="fileInput" name="file" accept=".pdf" style="display: none;">
        </div>

        <div class="file-info" id="fileInfo">
            <strong>📎 File đã chọn:</strong> <span id="fileName"></span>
            <br><small>Kích thước: <span id="fileSize"></span></small>
        </div>

        <button type="submit" class="btn-convert" id="convertBtn" disabled>
            🚀 Chuyển đổi sang Word
        </button>
    </form>

    <div class="loading" id="loading">
        <div class="spinner-border" role="status"></div>
        <p style="margin-top: 15px; color: #666;">Đang chuyển đổi, vui lòng đợi...</p>
    </div>

    <div class="result-card" id="resultCard">
        <h5>✅ Chuyển đổi thành công!</h5>
        <p>File Word của bạn đã sẵn sàng để tải về.</p>
        <a href="#" class="btn-download" id="downloadLink">⬇️ Tải file Word</a>
    </div>

    <div class="error-card" id="errorCard">
        <h5>❌ Lỗi</h5>
        <p id="errorMessage"></p>
    </div>

    <div class="features">
        <div class="feature">
            <span class="feature-icon">🔒</span>
            Bảo mật
        </div>
        <div class="feature">
            <span class="feature-icon">⚡</span>
            Nhanh chóng
        </div>
        <div class="feature">
            <span class="feature-icon">💰</span>
            Miễn phí
        </div>
        <div class="feature">
            <span class="feature-icon">🎯</span>
            Chính xác
        </div>
    </div>
</div>

<script>
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const convertBtn = document.getElementById('convertBtn');
    const uploadForm = document.getElementById('uploadForm');
    const loading = document.getElementById('loading');
    const resultCard = document.getElementById('resultCard');
    const errorCard = document.getElementById('errorCard');
    const errorMessage = document.getElementById('errorMessage');
    const downloadLink = document.getElementById('downloadLink');

    let selectedFile = null;

    // Click vào vùng upload
    uploadArea.addEventListener('click', () => fileInput.click());

    // Drag and drop
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    // Chọn file
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    // Xử lý file
    function handleFile(file) {
        if (!file.name.toLowerCase().endsWith('.pdf')) {
            showError('Chỉ chấp nhận file PDF');
            return;
        }
        if (file.size > 50 * 1024 * 1024) {
            showError('File vượt quá 50MB');
            return;
        }
        selectedFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        fileInfo.classList.add('show');
        convertBtn.disabled = false;
        resultCard.classList.remove('show');
        errorCard.classList.remove('show');
    }

    // Format kích thước file
    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
    }

    // Hiển thị lỗi
    function showError(message) {
        errorMessage.textContent = message;
        errorCard.classList.add('show');
        resultCard.classList.remove('show');
    }

    // Submit form
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!selectedFile) return;

        loading.classList.add('show');
        convertBtn.disabled = true;
        resultCard.classList.remove('show');
        errorCard.classList.remove('show');

        const formData = new FormData();
        formData.append('file', selectedFile);

        try {
            const response = await fetch('/convert', {
                method: 'POST',
                body: formData
            });
            const data = await response.json();

            loading.classList.remove('show');

            if (data.success) {
                downloadLink.href = '/download/' + data.filename;
                resultCard.classList.add('show');
            } else {
                showError(data.error || 'Có lỗi xảy ra trong quá trình chuyển đổi');
                convertBtn.disabled = false;
            }
        } catch (err) {
            loading.classList.remove('show');
            showError('Lỗi kết nối: ' + err.message);
            convertBtn.disabled = false;
        }
    });
</script>
</body>
</html>
'''

# ========== ROUTES ==========
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'Không tìm thấy file'})

        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'Chưa chọn file'})

        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Chỉ chấp nhận file PDF'})

        # Tạo tên file unique
        uid = uuid.uuid4().hex[:8]
        original_name = secure_filename(file.filename)
        base_name = os.path.splitext(original_name)[0]
        pdf_filename = f"{base_name}_{uid}.pdf"
        docx_filename = f"{base_name}_{uid}.docx"

        pdf_path = os.path.join(UPLOAD_DIR, pdf_filename)
        docx_path = os.path.join(OUTPUT_DIR, docx_filename)

        # Lưu file PDF
        file.save(pdf_path)

        # Chuyển đổi PDF sang Word
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()

        # Xóa file PDF gốc sau khi chuyển đổi
        try:
            os.remove(pdf_path)
        except Exception:
            pass

        return jsonify({
            'success': True,
            'filename': docx_filename,
            'message': 'Chuyển đổi thành công'
        })

    except Exception as e:
        print(f"❌ Lỗi chuyển đổi: {e}")
        return jsonify({'success': False, 'error': f'Lỗi chuyển đổi: {str(e)}'})

@app.route('/download/<filename>')
def download(filename):
    try:
        file_path = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(file_path):
            return "File không tồn tại", 404
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return f"Lỗi tải file: {str(e)}", 500

# ========== KHỞI ĐỘNG ==========
if __name__ == '__main__':
    print("=" * 60)
    print("🚀 PDF TO WORD CONVERTER")
    print("=" * 60)
    print("🌐 Mở trình duyệt: http://localhost:5000")
    print("📁 Thư mục upload:", UPLOAD_DIR)
    print("📁 Thư mục output:", OUTPUT_DIR)
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)