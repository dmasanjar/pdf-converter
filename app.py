import os
import subprocess
import traceback
import zipfile
import io
from flask import Flask, render_template, request, send_from_directory, send_file
from werkzeug.utils import secure_filename
from PIL import Image
from pdf2image import convert_from_path
from weasyprint import HTML

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {
    'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx',
    'jpg', 'jpeg', 'png', 'txt', 'html'
}

# Pastikan path LibreOffice sesuai instalasi di PC kamu
LIBREOFFICE_PATH = r'C:\\Program Files\\LibreOffice\\program\\soffice.exe'

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    uploaded_file = request.files['file']
    action = request.form.get('action')

    if not allowed_file(uploaded_file.filename):
        return {"results": [{"filename": uploaded_file.filename, "success": False, "error": "Unsupported file type."}]}

    filename = secure_filename(uploaded_file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    uploaded_file.save(input_path)

    ext = filename.rsplit('.', 1)[1].lower()
    output_filename = f"converted_{filename.rsplit('.', 1)[0]}.pdf"
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    try:
        if action == 'docx_to_pdf' and ext in ['doc', 'docx']:
            subprocess.run([LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf',
                           '--outdir', app.config['OUTPUT_FOLDER'], input_path], check=True)

        elif action == 'excel_to_pdf' and ext in ['xls', 'xlsx']:
            try:
                # Konversi ke HTML dulu
                html_filename = filename.rsplit('.', 1)[0] + '.html'
                html_path = os.path.join(app.config['OUTPUT_FOLDER'], html_filename)
                subprocess.run([LIBREOFFICE_PATH, '--headless', '--convert-to', 'html',
                               '--outdir', app.config['OUTPUT_FOLDER'], input_path], check=True)

                # Tambah CSS agar tabel rapi
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                css_style = """
                <style>
                  body { margin: 10px; font-family: Arial, sans-serif; font-size: 10pt; }
                  table { width: 100%; border-collapse: collapse; }
                  th, td { border: 1px solid #666; padding: 4px; text-align: left; font-size: 9pt; }
                  th { background: #eee; }
                </style>
                """
                html_content = html_content.replace('<head>', '<head>' + css_style)

                styled_html_path = html_path.replace('.html', '_styled.html')
                with open(styled_html_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)

                HTML(filename=styled_html_path).write_pdf(output_path)
                os.remove(html_path)
                os.remove(styled_html_path)
            except Exception:
                # fallback → langsung PDF
                subprocess.run([LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf',
                               '--outdir', app.config['OUTPUT_FOLDER'], input_path], check=True)

        elif action == 'powerpoint_to_pdf' and ext in ['ppt', 'pptx']:
            subprocess.run([LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf',
                           '--outdir', app.config['OUTPUT_FOLDER'], input_path], check=True)

        elif action == 'image_to_pdf' and ext in ['jpg', 'jpeg', 'png']:
            image = Image.open(input_path)
            if image.mode in ('RGBA', 'LA', 'P'):
                image = image.convert('RGB')
            image.save(output_path, 'PDF', resolution=100.0)

        elif action == 'text_to_pdf' and ext == 'txt':
            with open(input_path, 'r', encoding='utf-8') as f:
                text_content = f.read()
            HTML(string=f"<pre>{text_content}</pre>").write_pdf(output_path)

        elif action == 'html_to_pdf' and ext == 'html':
            HTML(input_path).write_pdf(output_path)

        elif action == 'pdf_to_image' and ext == 'pdf':
            images = convert_from_path(input_path)
            image_urls = []
            for i, img in enumerate(images):
                img_name = f"{filename}_page{i+1}.png"
                img_path = os.path.join(app.config['OUTPUT_FOLDER'], img_name)
                img.save(img_path, 'PNG')
                image_urls.append(f"/preview/{img_name}")
            os.remove(input_path)
            return {"images": image_urls}

        else:
            raise Exception("Unsupported file or conversion type.")

        os.remove(input_path)
        return {"results": [{"filename": filename, "success": True,
                             "preview_url": f"/preview/{output_filename}",
                             "download_filename": output_filename}]}

    except Exception as e:
        error_message = f"{str(e)}"
        print("❌ Conversion failed:", error_message)
        print(traceback.format_exc())
        return {"results": [{"filename": filename, "success": False, "error": error_message}]}


@app.route('/preview/<path:filename>')
def preview_file(filename):
    ext = filename.rsplit('.', 1)[1].lower()
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)

    if ext == 'pdf':
        return f"""
        <html><head><title>Preview {filename}</title></head>
        <body style='margin:0'>
            <iframe src='/output/{filename}' width='100%' height='100%' style='border:none;'></iframe>
            <div style='position:fixed;bottom:10px;right:10px;'>
                <a href='/output/{filename}' download='{filename}'>
                    <button style='padding:10px 20px;background:#4CAF50;color:#fff;border:none;border-radius:5px;'>Download</button>
                </a>
            </div>
        </body></html>
        """
    elif ext in ['png', 'jpg', 'jpeg']:
        return f"""
        <html><head><title>Preview {filename}</title></head>
        <body style='margin:0;text-align:center;'>
            <img src='/output/{filename}' style='max-width:90%;max-height:90vh;margin-top:20px;'/>
            <div style='margin-top:20px;'>
                <a href='/output/{filename}' download='{filename}'>
                    <button style='padding:10px 20px;background:#4CAF50;color:#fff;border:none;border-radius:5px;'>Download</button>
                </a>
            </div>
        </body></html>
        """
    else:
        return send_from_directory(app.config['OUTPUT_FOLDER'], filename)


@app.route('/output/<path:filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename)


@app.route('/download_selected', methods=['POST'])
def download_selected():
    try:
        files = request.json.get("files", [])
        if not files:
            return {"error": "No files selected"}, 400

        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w') as zf:
            for fname in files:
                file_path = os.path.join(app.config['OUTPUT_FOLDER'], fname)
                if os.path.isfile(file_path):
                    zf.write(file_path, arcname=fname)
        memory_file.seek(0)

        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name='converted_files.zip'
        )
    except Exception as e:
        return {"error": str(e)}, 500


if __name__ == '__main__':
    app.run(debug=True, port=5050)
