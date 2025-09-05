from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, send_file, after_this_request
import os
import subprocess
import zipfile
import io
import aspose.slides as slides
import jpype.dbapi2
from pdf2docx import Converter
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from PIL import Image
from pdf2image import convert_from_path
import img2pdf
import cv2
import numpy as np
from skimage.filters import threshold_local
import pdfkit
from flask import make_response
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch

app = Flask(__name__)
app.secret_key = 'your_super_secret_key_change_this'

# --- CONFIGURATION ---
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# --- HELPER & CLEANUP FUNCTIONS ---
def get_human_readable_size(size_in_bytes):
    if size_in_bytes is None: return ""
    power = 1024; n = 0
    power_labels = {0: '', 1: 'KB', 2: 'MB', 3: 'GB', 4: 'TB'}
    while size_in_bytes >= power and n < len(power_labels):
        size_in_bytes /= power; n += 1
    return f"{size_in_bytes:.2f} {power_labels[n]}"

def cleanup_files():
    original_files = session.pop('original_files', [])
    processed_files = session.pop('processed_files', [])
    for filename in original_files:
        try: os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        except (FileNotFoundError, TypeError, OSError): pass
    for filename in processed_files:
        try: os.remove(os.path.join(app.config['PROCESSED_FOLDER'], filename))
        except (FileNotFoundError, TypeError, OSError): pass

# --- PAGE RENDERING ROUTES ---
@app.route('/')
def index(): return render_template('index.html')
@app.route('/word-to-pdf-page')
def word_to_pdf_page(): return render_template('word_to_pdf.html')
@app.route('/pdf-to-word-page')
def pdf_to_word_page(): return render_template('pdf_to_word.html')
@app.route('/compress-pdf-page')
def compress_pdf_page(): return render_template('compress.html')
@app.route('/merge-pdf-page')
def merge_pdf_page(): return render_template('merge.html')
@app.route('/split-pdf-page')
def split_pdf_page(): return render_template('split.html')
@app.route('/protect-pdf-page')
def protect_pdf_page(): return render_template('protect.html')
@app.route('/unlock-pdf-page')
def unlock_pdf_page(): return render_template('unlock.html')
@app.route('/rotate-pdf-page')
def rotate_pdf_page(): return render_template('rotate.html')
@app.route('/pdf-to-jpg-page')
def pdf_to_jpg_page(): return render_template('pdf_to_jpg.html')
@app.route('/jpg-to-pdf-page')
def jpg_to_pdf_page(): return render_template('jpg_to_pdf.html')
@app.route('/pdf-to-png-page')
def pdf_to_png_page(): return render_template('pdf_to_png.html')
@app.route('/png-to-pdf-page')
def png_to_pdf_page(): return render_template('png_to_pdf.html')
@app.route('/scan-document-page')
def scan_document_page(): return render_template('scan.html')
@app.route('/pptx-to-pdf-page')
def pptx_to_pdf_page(): return render_template('pptx_to_pdf.html')
@app.route('/add-page-numbers-page')
def add_page_numbers_page(): return render_template('add_page_numbers.html')
@app.route('/png-jpg-tools')
def png_jpg_tools_page(): return render_template('png_jpg_tools_page.html')

@app.route('/downloads')
def download_page():
    filenames = session.get('processed_files', [])
    files_with_sizes = []
    for name in filenames:
        filepath = os.path.join(app.config['PROCESSED_FOLDER'], name)
        try:
            size_bytes = os.path.getsize(filepath)
            size_readable = get_human_readable_size(size_bytes)
            files_with_sizes.append({'name': name, 'size': size_readable})
        except FileNotFoundError: continue
    return render_template('downloads.html', files=files_with_sizes)

# --- FILE PROCESSING ROUTES ---
@app.route('/pptx-to-pdf', methods=['GET', 'POST'])
def pptx_to_pdf():
    if request.method == 'POST':
        if 'file' not in request.files or request.files['file'].filename == '':
            flash({'title': 'No File Selected', 'message': 'Please select a PowerPoint file (.ppt or .pptx) to upload.'}, 'warning')
            return redirect(request.url)

        file = request.files['file']

        if file and (file.filename.endswith('.ppt') or file.filename.endswith('.pptx')):
            try:
                original_filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                output_filename = os.path.splitext(file.filename)[0] + '.pdf'
                output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)

                # Save the uploaded file
                file.save(original_filepath)

                # --- NEW, RELIABLE CONVERSION LOGIC ---
                # Load the presentation
                pres = slides.Presentation(original_filepath)
                
                # Save it directly to PDF
                pres.save(output_filepath, slides.export.SaveFormat.PDF)
                # --- END OF NEW LOGIC ---

                session['original_files'] = []
                session['processed_files'] = [output_filename]
                return redirect(url_for('download_page'))

            except Exception as e:
                print(f"ASPOSE CONVERSION ERROR: {e}")
                flash({
                    'title': 'Conversion Failed',
                    'message': "We couldn't process this file. It might be an unsupported format or corrupted.",
                    'suggestion': "Please try saving the presentation again and re-uploading."
                }, 'danger')
                return redirect(request.url)
        else:
            flash({'title': 'Invalid File Type', 'message': 'Please upload a valid .ppt or .pptx file.'}, 'warning')
            return redirect(request.url)
            
    return render_template('pptx_to_pdf.html')

@app.route('/add-page-numbers', methods=['POST'])
def add_page_numbers():
    file = request.files.get('file')
    position = request.form.get('position', 'bottom-center')
    num_format = request.form.get('format', '1')
    if not file: return "Missing file.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    reader = PdfReader(filepath)
    writer = PdfWriter()
    if reader.is_encrypted:
        os.remove(filepath)
        return "Cannot add page numbers to an encrypted file. Please unlock it first.", 400
    total_pages = len(reader.pages)
    for i, page in enumerate(reader.pages):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(page.mediabox.width, page.mediabox.height))
        page_num_text = ""
        if num_format == '1': page_num_text = f"{i + 1}"
        elif num_format == '1_of_n': page_num_text = f"Page {i + 1} of {total_pages}"
        elif num_format == '- 1 -': page_num_text = f"- {i + 1} -"
        page_width, page_height = float(page.mediabox.width), float(page.mediabox.height)
        y = 0
        if 'bottom' in position: y = 0.5 * inch
        if 'top' in position: y = page_height - 0.5 * inch
        if 'center' in position: can.drawCentredString(page_width / 2, y, page_num_text)
        elif 'right' in position: can.drawString(page_width - 1 * inch, y, page_num_text)
        elif 'left' in position: can.drawString(1 * inch, y, page_num_text)
        can.save()
        packet.seek(0)
        number_pdf = PdfReader(packet)
        page.merge_page(number_pdf.pages[0])
        writer.add_page(page)
    output_filename = f"numbered_{original_filename}"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: writer.write(f)
    session['original_files'] = [original_filename]
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/pdf-to-png', methods=['POST'])
def pdf_to_png():
    file = request.files.get('file')
    if not file: return "Missing file.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    poppler_path = r"C:\Release-25.07.0-0\poppler-25.07.0\Library\bin"
    images = convert_from_path(filepath, poppler_path=poppler_path, fmt='png')
    processed_files = []
    base_filename = os.path.splitext(original_filename)[0]
    for i, image in enumerate(images):
        png_filename = f"{base_filename}_page_{i + 1}.png"
        png_filepath = os.path.join(app.config['PROCESSED_FOLDER'], png_filename)
        image.save(png_filepath, 'PNG')
        processed_files.append(png_filename)
    session['original_files'] = [original_filename]
    session['processed_files'] = processed_files
    return redirect(url_for('download_page'))

@app.route('/png-to-pdf', methods=['POST'])
def png_to_pdf():
    files = request.files.getlist('files')
    if not files: return "No images uploaded.", 400
    image_paths = []
    original_filenames = []
    for file in files:
        if file.filename:
            original_filenames.append(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            image_paths.append(filepath)
    output_filename = "converted_images_from_png.pdf"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: f.write(img2pdf.convert(image_paths))
    session['original_files'] = original_filenames
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/protect-pdf', methods=['POST'])
def protect_pdf():
    file = request.files.get('file')
    password = request.form.get('password')
    if not file or not password: return "Missing file or password.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    reader = PdfReader(filepath)
    writer = PdfWriter()
    for page in reader.pages: writer.add_page(page)
    writer.encrypt(password)
    output_filename = f"protected_{original_filename}"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: writer.write(f)
    session['original_files'] = [original_filename]
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/unlock-pdf', methods=['POST'])
def unlock_pdf():
    file = request.files.get('file')
    password = request.form.get('password')
    if not file: return "Missing file.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    reader = PdfReader(filepath)
    writer = PdfWriter()
    if reader.is_encrypted:
        if not reader.decrypt(password): return "Incorrect password.", 400
    for page in reader.pages: writer.add_page(page)
    output_filename = f"unlocked_{original_filename}"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: writer.write(f)
    session['original_files'] = [original_filename]
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/rotate-pdf', methods=['POST'])
def rotate_pdf():
    file = request.files.get('file')
    rotation = int(request.form.get('rotation', 90))
    if not file: return "Missing file.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    reader = PdfReader(filepath)
    writer = PdfWriter()
    if reader.is_encrypted:
        os.remove(filepath) 
        return "This PDF is password-protected. Please use the 'Unlock PDF' tool first.", 400
    for page in reader.pages:
        page.rotate(rotation)
        writer.add_page(page)
    output_filename = f"rotated_{original_filename}"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: writer.write(f)
    session['original_files'] = [original_filename]
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/pdf-to-jpg', methods=['POST'])
def pdf_to_jpg():
    file = request.files.get('file')
    if not file: return "Missing file.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    poppler_path = r"C:\Release-25.07.0-0\poppler-25.07.0\Library\bin"
    images = convert_from_path(filepath, poppler_path=poppler_path)
    processed_files = []
    base_filename = os.path.splitext(original_filename)[0]
    for i, image in enumerate(images):
        jpg_filename = f"{base_filename}_page_{i + 1}.jpg"
        jpg_filepath = os.path.join(app.config['PROCESSED_FOLDER'], jpg_filename)
        image.save(jpg_filepath, 'JPEG')
        processed_files.append(jpg_filename)
    session['original_files'] = [original_filename]
    session['processed_files'] = processed_files
    return redirect(url_for('download_page'))

@app.route('/jpg-to-pdf', methods=['POST'])
def jpg_to_pdf():
    files = request.files.getlist('files')
    if not files: return "No images uploaded.", 400
    image_paths = []
    original_filenames = []
    for file in files:
        if file.filename:
            original_filenames.append(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            image_paths.append(filepath)
    output_filename = "converted_images.pdf"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: f.write(img2pdf.convert(image_paths))
    session['original_files'] = original_filenames
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/split-pdf', methods=['POST'])
def split_pdf():
    file = request.files.get('file')
    page_ranges_str = request.form.get('page_ranges')
    if not file or not page_ranges_str: return "Missing file or page ranges.", 400
    original_filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
    file.save(filepath)
    reader = PdfReader(filepath)
    writer = PdfWriter()
    try:
        pages_to_extract = set()
        for part in page_ranges_str.split(','):
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                for i in range(start, end + 1): pages_to_extract.add(i)
            else:
                pages_to_extract.add(int(part))
        for page_num in sorted(list(pages_to_extract)):
            if 1 <= page_num <= len(reader.pages): writer.add_page(reader.pages[page_num - 1])
    except ValueError:
        return "Invalid page range format.", 400
    output_filename = f"split_{original_filename}"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    with open(output_filepath, "wb") as f: writer.write(f)
    session['original_files'] = [original_filename]
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/merge-pdf', methods=['POST'])
def merge_pdf():
    files = request.files.getlist('files')
    order_str = request.form.get('file_order')
    ordered_filenames = order_str.split(',') if order_str else []
    original_filenames = [f.filename for f in files if f.filename]
    for file in files:
        if file.filename: file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
    merger = PdfMerger()
    for filename in ordered_filenames:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(filepath): merger.append(filepath)
    output_filename = "merged_document.pdf"
    output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
    if os.path.exists(output_filepath): os.remove(output_filepath)
    merger.write(output_filepath)
    merger.close()
    session['original_files'] = original_filenames
    session['processed_files'] = [output_filename]
    return redirect(url_for('download_page'))

@app.route('/word-to-pdf', methods=['POST'])
def convert_word_to_pdf():
    files = request.files.getlist('files')
    processed_files = []
    output_dir = app.config['PROCESSED_FOLDER']
    session['original_files'] = [f.filename for f in files if f.filename]
    for file in files:
        if file.filename and file.filename.endswith(('.doc', '.docx')):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            soffice_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
            subprocess.run([soffice_path, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, filepath])
            pdf_filename = os.path.splitext(file.filename)[0] + '.pdf'
            processed_files.append(pdf_filename)
    session['processed_files'] = processed_files
    return redirect(url_for('download_page'))

@app.route('/pdf-to-word', methods=['POST'])
def convert_pdf_to_word():
    files = request.files.getlist('files')
    processed_files = []
    session['original_files'] = [f.filename for f in files if f.filename]
    for file in files:
        if file.filename and file.filename.endswith('.pdf'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            docx_filename = os.path.splitext(file.filename)[0] + '.docx'
            docx_filepath = os.path.join(app.config['PROCESSED_FOLDER'], docx_filename)
            cv = Converter(filepath)
            cv.convert(docx_filepath, start=0, end=None)
            cv.close()
            processed_files.append(docx_filename)
    session['processed_files'] = processed_files
    return redirect(url_for('download_page'))

@app.route('/compress-pdf', methods=['POST'])
def compress_pdf():
    files = request.files.getlist('files')
    try:
        target_mb = float(request.form.get('target_size'))
        target_bytes = target_mb * 1024 * 1024
    except (ValueError, TypeError):
        target_bytes = 2 * 1024 * 1024
    processed_files = []
    session['original_files'] = [f.filename for f in files if f.filename]
    quality_levels = {'/prepress': '-dPDFSETTINGS=/prepress', '/printer': '-dPDFSETTINGS=/printer', '/ebook': '-dPDFSETTINGS=/ebook', '/screen': '-dPDFSETTINGS=/screen'}
    for file in files:
        if file.filename and file.filename.endswith('.pdf'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            final_filename = ""
            for quality_name, setting in quality_levels.items():
                temp_filename = f"temp_{quality_name.replace('/', '')}_{file.filename}"
                output_filepath = os.path.join(app.config['PROCESSED_FOLDER'], temp_filename)
                subprocess.run(['gswin64c', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4', setting, '-dNOPAUSE', '-dQUIET', '-dBATCH', f'-sOutputFile={output_filepath}', filepath])
                if not os.path.exists(output_filepath): continue
                output_size = os.path.getsize(output_filepath)
                if output_size <= target_bytes:
                    final_filename = f"compressed_{file.filename}"
                    final_filepath = os.path.join(app.config['PROCESSED_FOLDER'], final_filename)
                    if os.path.exists(final_filepath): os.remove(final_filepath)
                    os.rename(output_filepath, final_filepath)
                    break
                else:
                    os.remove(output_filepath)
            if not final_filename:
                final_filename = f"compressed_{file.filename}"
                final_filepath = os.path.join(app.config['PROCESSED_FOLDER'], final_filename)
                if os.path.exists(final_filepath): os.remove(final_filepath)
                subprocess.run(['gswin64c', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4', quality_levels['/screen'], '-dNOPAUSE', '-dQUIET', '-dBATCH', f'-sOutputFile={final_filepath}', filepath])
            processed_files.append(final_filename)
    session['processed_files'] = processed_files
    return redirect(url_for('download_page'))

@app.route('/png-to-jpg', methods=['POST'])
def png_to_jpg():
    files = request.files.getlist('files')
    memory_file = io.BytesIO()

    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in files:
            if file and file.filename.endswith('.png'):
                # Read the image and convert it
                img = Image.open(file.stream).convert('RGB')
                
                # Create a buffer for the new JPG image
                output_buffer = io.BytesIO()
                img.save(output_buffer, 'JPEG')
                output_buffer.seek(0)
                
                # Write the new JPG to the zip file
                new_filename = os.path.splitext(file.filename)[0] + '.jpg'
                zf.writestr(new_filename, output_buffer.read())

    memory_file.seek(0)
    return send_file(
        memory_file,
        download_name='converted_jpgs.zip',
        as_attachment=True,
        mimetype='application/zip'
    )

# Route to handle JPG to PNG conversion
@app.route('/jpg-to-png', methods=['POST'])
def jpg_to_png():
    files = request.files.getlist('files')
    memory_file = io.BytesIO()

    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in files:
            if file and (file.filename.endswith('.jpg') or file.filename.endswith('.jpeg')):
                # Read the image
                img = Image.open(file.stream)
                
                # Create a buffer for the new PNG image
                output_buffer = io.BytesIO()
                img.save(output_buffer, 'PNG')
                output_buffer.seek(0)
                
                # Write the new PNG to the zip file
                new_filename = os.path.splitext(file.filename)[0] + '.png'
                zf.writestr(new_filename, output_buffer.read())

    memory_file.seek(0)
    return send_file(
        memory_file,
        download_name='converted_pngs.zip',
        as_attachment=True,
        mimetype='application/zip'
    )

# --- FILE DOWNLOAD ROUTES ---
@app.route('/download/<filename>')
def download_file(filename):
    @after_this_request
    def cleanup(response):
        cleanup_files()
        return response
    return send_from_directory(app.config['PROCESSED_FOLDER'], filename, as_attachment=True)

@app.route('/download-zip')
def download_zip():
    files = session.get('processed_files', [])
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for filename in files:
            filepath = os.path.join(app.config['PROCESSED_FOLDER'], filename)
            if os.path.exists(filepath): zf.write(filepath, arcname=filename)
    memory_file.seek(0)
    @after_this_request
    def cleanup(response):
        cleanup_files()
        return response
    return send_file(memory_file, download_name='ZenPDF_Files.zip', as_attachment=True, mimetype='application/zip')

if __name__ == '__main__':
    app.run(debug=True)
