import os
from flask import Flask, render_template, request, jsonify, send_from_directory, url_for, send_file
from werkzeug.utils import secure_filename
import uuid
import zipfile
import shutil
from concurrent.futures import ThreadPoolExecutor
import threading
import logging
import tempfile
import json
import sys

# Global imports with error handling for Vercel stability
GLOBAL_IMPORT_ERROR = None
try:
    import pandas as pd
    import automate_invoices as ai
except Exception as e:
    import traceback
    GLOBAL_IMPORT_ERROR = f"{str(e)}\n{traceback.format_exc()}"

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Explicitly use /tmp on Linux/Vercel
base_temp = '/tmp' if sys.platform.startswith('linux') else tempfile.gettempdir()

app.config['UPLOAD_FOLDER'] = os.path.join(base_temp, 'fly91_uploads')
app.config['TEMP_OUTPUT'] = os.path.join(base_temp, 'fly91_temp_output')
app.config['REPO_MEDIA_FOLDER'] = os.path.join(os.path.dirname(__file__), 'media')
app.config['UPLOADED_MEDIA_FOLDER'] = os.path.join(base_temp, 'fly91_media')
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
    if GLOBAL_IMPORT_ERROR:
        return jsonify({'error': f"Server Startup Error (Imports Failed): {GLOBAL_IMPORT_ERROR}"}), 500
        
    if 'excel' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['excel']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename, ALLOWED_EXTENSIONS_EXCEL):
        filename = secure_filename(file.filename)
        unique_filename = f"{uuid.uuid4().hex}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        
        try:
            # Ensure directories exist exactly before saving
            os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
            file.save(filepath)
            
            df = ai.get_excel_data_rows(filepath)
            if df.empty:
                return jsonify({'error': 'Excel file contains no data rows (Invoicenumber column must not be empty)'}), 400
                
            # Cache data in memory so we can delete the file immediately
            cached_rows = df.to_json(orient='records')
            first_row = {
                'index': 0,
                'invoice_no': str(df.iloc[0].get('Invoicenumber', 'N/A')),
                'customer': str(df.iloc[0].get('Customer Name ', 'N/A')),
                'pnr': str(df.iloc[0].get('PNRNumber', 'N/A'))
            }
            # Store full data keyed by session token so we never need the file again
            with progress_lock:
                batch_progress[unique_filename] = {
                    'cached_df': cached_rows,
                    'excel_path': filepath  # kept only for address sheet lookup
                }
            return jsonify({
                'success': True,
                'first_row': first_row,
                'total_rows': len(df),
                'filename': unique_filename
            })
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            logger.error(f"Upload process failed: {error_details}")
            print(f"DEBUG: Upload error: {e}") # Visible in local terminal
            if os.path.exists(filepath): os.remove(filepath)
            return jsonify({'error': f"Processing Error: {str(e)}"}), 500
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

def _cleanup_old_previews():
    """Delete all preview_*.pdf files older than 5 minutes from TEMP_OUTPUT."""
    import time
    temp_dir = app.config['TEMP_OUTPUT']
    now = time.time()
    try:
        for f in os.listdir(temp_dir):
            if f.startswith('preview_') and f.endswith('.pdf'):
                fp = os.path.join(temp_dir, f)
                try:
                    if now - os.path.getmtime(fp) > 300:
                        os.remove(fp)
                except Exception:
                    pass
    except Exception:
        pass

@app.route('/preview_first')
def preview_first():
    excel_filename = request.args.get('excel')
    if not excel_filename:
        return "Missing excel filename", 400

    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    if not os.path.exists(excel_path):
        return "Excel file not found", 404

    # Optional pre-saved positions for auto-align
    seal_pos = request.args.get('seal_pos')
    sign_pos = request.args.get('sign_pos')
    try:
        seal_pos_obj = json.loads(seal_pos) if seal_pos else None
        sign_pos_obj = json.loads(sign_pos) if sign_pos else None
    except:
        seal_pos_obj = sign_pos_obj = None

    _cleanup_old_previews()
    import automate_invoices as ai
    df = ai.get_excel_data_rows(excel_path)
    data = ai.get_invoicing_data(df, 0, excel_path)

    pdf_filename = f"preview_{uuid.uuid4().hex}.pdf"
    pdf_path = os.path.join(app.config['TEMP_OUTPUT'], pdf_filename)

    ai.generate_kind_pdf(data, pdf_path, seal_pos=seal_pos_obj, sign_pos=sign_pos_obj)

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
    if not os.path.exists(excel_path):
        return jsonify({'error': 'Excel file not found'}), 404

    _cleanup_old_previews()
    import automate_invoices as ai
    df = ai.get_excel_data_rows(excel_path)
    data = ai.get_invoicing_data(df, 0, excel_path)

    pdf_filename = f"preview_{uuid.uuid4().hex}.pdf"
    pdf_path = os.path.join(app.config['TEMP_OUTPUT'], pdf_filename)

    ai.generate_kind_pdf(data, pdf_path, seal_pos=seal_pos, sign_pos=sign_pos)

    return jsonify({'success': True, 'pdf_url': url_for('get_temp_pdf', filename=pdf_filename)})

@app.route('/temp_pdf/<filename>')
def get_temp_pdf(filename):
    file_path = os.path.join(app.config['TEMP_OUTPUT'], filename)
    if not os.path.exists(file_path):
        return "File not found", 404
    # Stream the file then delete it if it is a zip (already downloaded)
    def _send_and_delete():
        response = send_from_directory(app.config['TEMP_OUTPUT'], filename)
        if filename.endswith('.zip'):
            @response.call_on_close
            def _del():
                try:
                    os.remove(file_path)
                    logger.info(f"Deleted temp zip: {file_path}")
                except Exception:
                    pass
        return response
    return _send_and_delete()

@app.route('/batch_progress/<session_id>')
def get_batch_progress(session_id):
    with progress_lock:
        return jsonify(batch_progress.get(session_id, {'current': 0, 'total': 0, 'status': 'unknown'}))

def process_single_pdf(i, df, excel_path, session_dir, seal_pos, sign_pos, session_id):
    try:
        import automate_invoices as ai
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
            import automate_invoices as ai
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

            # Clean up session PDFs and the uploaded excel immediately
            shutil.rmtree(session_dir, ignore_errors=True)
            if os.path.exists(excel_path):
                os.remove(excel_path)
                logger.info(f"Deleted uploaded excel: {excel_path}")

            with progress_lock:
                batch_progress[session_id]['status'] = 'completed'
                batch_progress[session_id]['zip_url'] = f"/temp_pdf/{zip_filename}"
                # zip itself is deleted on download via get_temp_pdf
        except Exception as e:
            logger.error(f"Background batch error: {e}")
            # Clean up on failure too
            shutil.rmtree(session_dir, ignore_errors=True)
            if os.path.exists(excel_path):
                try: os.remove(excel_path)
                except Exception: pass
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
