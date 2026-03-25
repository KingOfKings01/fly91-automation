import os
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory, url_for, send_file
from werkzeug.utils import secure_filename
import automate_invoices as ai
import uuid
import zipfile
import shutil
from concurrent.futures import ThreadPoolExecutor
import threading

import logging
import tempfile

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
# Use /tmp for Vercel deployment (read-only filesystem elsewhere)
app.config['UPLOAD_FOLDER'] = os.path.join(tempfile.gettempdir(), 'fly91_uploads')
app.config['TEMP_OUTPUT'] = os.path.join(tempfile.gettempdir(), 'fly91_temp_output')
app.config['REPO_MEDIA_FOLDER'] = os.path.join(os.path.dirname(__file__), 'media')
app.config['UPLOADED_MEDIA_FOLDER'] = os.path.join(tempfile.gettempdir(), 'fly91_media')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

def ensure_dirs():
    """Ensure all required directories exist. Called before first file operation."""
    for path in [app.config['UPLOAD_FOLDER'], app.config['TEMP_OUTPUT'], app.config['UPLOADED_MEDIA_FOLDER']]:
        if not os.path.exists(path):
            try:
                os.makedirs(path, exist_ok=True)
                logger.info(f"Created directory: {path}")
            except Exception as e:
                logger.error(f"Failed to create {path}: {e}")

@app.route('/health')
def health():
    return "OK", 200

@app.before_request
def before_request():
    ensure_dirs()

ALLOWED_EXTENSIONS_EXCEL = {'xlsx', 'xlsm', 'xls'}
ALLOWED_EXTENSIONS_IMG = {'png', 'jpg', 'jpeg'}

# Progress tracking
batch_progress = {}
progress_lock = threading.Lock()
def allowed_file(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'excel' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['excel']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS_EXCEL):
        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4().hex}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(filepath)
        
        try:
            df = ai.get_excel_data_rows(filepath)
            first_row = {
                'index': 0,
                'invoice_no': str(df.iloc[0].get('Invoicenumber', 'N/A')),
                'customer': str(df.iloc[0].get('Customer Name ', 'N/A')),
                'pnr': str(df.iloc[0].get('PNRNumber', 'N/A'))
            }
            return jsonify({
                'success': True, 
                'first_row': first_row, 
                'total_rows': len(df),
                'filename': unique_filename
            })
        except Exception as e:
            if os.path.exists(filepath): os.remove(filepath)
            return jsonify({'error': str(e)}), 500
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/upload_media', methods=['POST'])
def upload_media():
    mtype = request.form.get('type')
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS_IMG):
        if mtype == 'seal':
            target_name = 'image2.png'
        elif mtype == 'sign':
            target_name = 'image3.png'
        else:
            return jsonify({'error': 'Invalid type'}), 400
            
        filepath = os.path.join(app.config['UPLOADED_MEDIA_FOLDER'], target_name)
        file.save(filepath)
        return jsonify({'success': True, 'url': url_for('get_media', filename=target_name, _t=os.path.getmtime(filepath))})
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/media/<filename>')
def get_media(filename):
    # Check /tmp first, then repo defaults
    tmp_path = os.path.join(app.config['UPLOADED_MEDIA_FOLDER'], filename)
    if os.path.exists(tmp_path):
        return send_from_directory(app.config['UPLOADED_MEDIA_FOLDER'], filename)
    return send_from_directory(app.config['REPO_MEDIA_FOLDER'], filename)

@app.route('/preview_first')
def preview_first():
    excel_filename = request.args.get('excel')
    if not excel_filename:
        return "Missing excel filename", 400
    
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    if not os.path.exists(excel_path):
        return "Excel file not found", 404
        
    df = ai.get_excel_data_rows(excel_path)
    data = ai.get_invoicing_data(df, 0, excel_path)
    
    pdf_filename = f"preview_{uuid.uuid4().hex}.pdf"
    pdf_path = os.path.join(app.config['TEMP_OUTPUT'], pdf_filename)
    
    ai.generate_kind_pdf(data, pdf_path)
    
    return render_template('preview.html', 
                           pdf_url=url_for('get_temp_pdf', filename=pdf_filename), 
                           excel_filename=excel_filename,
                           invoice_no=data['invoice_no'])

@app.route('/refresh_preview', methods=['POST'])
def refresh_preview():
    req_data = request.json
    excel_filename = req_data.get('excel_filename')
    seal_pos = req_data.get('seal_pos')
    sign_pos = req_data.get('sign_pos')
    
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    df = ai.get_excel_data_rows(excel_path)
    data = ai.get_invoicing_data(df, 0, excel_path)
    
    pdf_filename = f"preview_{uuid.uuid4().hex}.pdf"
    pdf_path = os.path.join(app.config['TEMP_OUTPUT'], pdf_filename)
    
    ai.generate_kind_pdf(data, pdf_path, seal_pos=seal_pos, sign_pos=sign_pos)
    
    return jsonify({'success': True, 'pdf_url': url_for('get_temp_pdf', filename=pdf_filename)})

@app.route('/temp_pdf/<filename>')
def get_temp_pdf(filename):
    return send_from_directory(app.config['TEMP_OUTPUT'], filename)

@app.route('/batch_progress/<session_id>')
def get_batch_progress(session_id):
    with progress_lock:
        return jsonify(batch_progress.get(session_id, {'current': 0, 'total': 0, 'status': 'unknown'}))

def process_single_pdf(i, df, excel_path, session_dir, seal_pos, sign_pos, session_id):
    try:
        data = ai.get_invoicing_data(df, i, excel_path)
        if not data['invoice_no'] or data['invoice_no'] == 'nan': 
            with progress_lock:
                batch_progress[session_id]['current'] += 1
            return
        
        bifurcation = data.get('folder_bifurcation', 'Unknown')
        target_dir = os.path.join(session_dir, bifurcation)
        os.makedirs(target_dir, exist_ok=True)
        
        safe_inv = "".join([c for c in data['invoice_no'] if c.isalnum() or c in (' ', '-', '_')]).strip()
        pdf_name = f"{safe_inv}.pdf"
        pdf_path = os.path.join(target_dir, pdf_name)
        
        ai.generate_kind_pdf(data, pdf_path, seal_pos=seal_pos, sign_pos=sign_pos)
        with progress_lock:
            batch_progress[session_id]['current'] += 1
    except Exception as e:
        print(f"Error processing row {i}: {e}")
        with progress_lock:
            batch_progress[session_id]['current'] += 1

def run_background_batch(session_id, excel_path, session_dir, seal_pos, sign_pos):
    with app.app_context():
        try:
            df = ai.get_excel_data_rows(excel_path)
            with progress_lock:
                batch_progress[session_id] = {'current': 0, 'total': len(df), 'status': 'processing'}
            
            with ThreadPoolExecutor(max_workers=10) as executor:
                for i in range(len(df)):
                    executor.submit(process_single_pdf, i, df, excel_path, session_dir, seal_pos, sign_pos, session_id)
            
            zip_filename = f"Invoices_{session_id}.zip"
            zip_path = os.path.join(app.config['TEMP_OUTPUT'], zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for root, dirs, files in os.walk(session_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, session_dir)
                        zipf.write(file_path, arcname)
            
            shutil.rmtree(session_dir)
            if os.path.exists(excel_path): os.remove(excel_path)
            
            with progress_lock:
                batch_progress[session_id]['status'] = 'completed'
                batch_progress[session_id]['zip_url'] = f"/temp_pdf/{zip_filename}"
        except Exception as e:
            print(f"Background batch error: {e}")
            with progress_lock:
                batch_progress[session_id]['status'] = 'error'

@app.route('/generate_batch', methods=['POST'])
def generate_batch():
    req_data = request.json
    excel_filename = req_data.get('excel_filename')
    seal_pos = req_data.get('seal_pos')
    sign_pos = req_data.get('sign_pos')
    
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    if not os.path.exists(excel_path):
        return jsonify({'error': 'Excel file missing'}), 404
        
    session_id = uuid.uuid4().hex
    session_dir = os.path.join(app.config['TEMP_OUTPUT'], session_id)
    os.makedirs(session_dir, exist_ok=True)
    
    # Initialize progress
    with progress_lock:
        batch_progress[session_id] = {'current': 0, 'total': 0, 'status': 'starting'}
    
    # Start background thread
    thread = threading.Thread(target=run_background_batch, args=(session_id, excel_path, session_dir, seal_pos, sign_pos))
    thread.start()
    
    return jsonify({'success': True, 'session_id': session_id})

if __name__ == '__main__':
    app.run(debug=True, port=5000, threaded=True)
