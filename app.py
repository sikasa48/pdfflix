from flask import Flask, render_template, request, send_file, jsonify
import os, sys, uuid, subprocess, io, tempfile
from werkzeug.utils import secure_filename
from pypdf import PdfWriter, PdfReader

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

ALLOWED_WORD = {'doc', 'docx'}
ALLOWED_PDF  = {'pdf'}

def allowed_file(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

# ─────────────────────────────────────────────
#  CONVERSION : Windows (PowerShell+Word)
#               Linux/Mac (LibreOffice)
# ─────────────────────────────────────────────
def convert_word_to_pdf_bytes(word_path):
    abs_input = os.path.abspath(word_path)

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        tmp_pdf = os.path.abspath(tmp.name)

    try:
        if sys.platform == 'win32':
            # ── Windows : PowerShell pilote Microsoft Word ──
            ps = (
                "$w = New-Object -ComObject Word.Application;"
                "$w.Visible = $false;"
                "$w.DisplayAlerts = 0;"
                f"$d = $w.Documents.Open('{abs_input}');"
                f"$d.SaveAs2('{tmp_pdf}', 17);"
                "$d.Close($false);"
                "$w.Quit();"
            )
            r = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                capture_output=True, text=True, timeout=120
            )
            if not os.path.exists(tmp_pdf) or os.path.getsize(tmp_pdf) == 0:
                raise Exception(f"Erreur Word/PowerShell : {r.stderr or r.stdout or 'inconnu'}")

        else:
            # ── Linux / Mac : LibreOffice headless ──
            out_dir = os.path.dirname(tmp_pdf)
            r = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'pdf',
                 '--outdir', out_dir, abs_input],
                capture_output=True, text=True, timeout=120
            )
            base   = os.path.splitext(os.path.basename(abs_input))[0]
            lo_out = os.path.join(out_dir, base + '.pdf')
            if os.path.exists(lo_out):
                os.replace(lo_out, tmp_pdf)
            if not os.path.exists(tmp_pdf) or os.path.getsize(tmp_pdf) == 0:
                raise Exception(f"Erreur LibreOffice : {r.stderr or 'inconnu'}")

        with open(tmp_pdf, 'rb') as f:
            return f.read()

    finally:
        if os.path.exists(tmp_pdf):
            os.remove(tmp_pdf)


# ─────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_route():
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier sélectionné'}), 400

    file = request.files['file']
    if not allowed_file(file.filename, ALLOWED_WORD):
        return jsonify({'error': 'Format non supporté. Utilisez .doc ou .docx'}), 400

    original_name = secure_filename(file.filename)
    base_name     = os.path.splitext(original_name)[0]
    ext           = os.path.splitext(original_name)[1]

    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp_path = tmp.name
        file.save(tmp_path)

    try:
        pdf_bytes = convert_word_to_pdf_bytes(tmp_path)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    return send_file(
        io.BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=f"{base_name}.pdf"
    )


@app.route('/merge', methods=['POST'])
def merge_pdfs():
    if 'files' not in request.files:
        return jsonify({'error': 'Aucun fichier sélectionné'}), 400

    files = request.files.getlist('files')
    if len(files) < 2:
        return jsonify({'error': 'Sélectionnez au moins 2 fichiers PDF'}), 400

    tmp_paths = []
    try:
        for f in files:
            if not allowed_file(f.filename, ALLOWED_PDF):
                return jsonify({'error': f'"{f.filename}" n\'est pas un PDF valide'}), 400
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                f.save(tmp.name)
                tmp_paths.append(tmp.name)

        writer      = PdfWriter()
        total_pages = 0
        for path in tmp_paths:
            reader = PdfReader(path)
            for page in reader.pages:
                writer.add_page(page)
            total_pages += len(reader.pages)

        buffer = io.BytesIO()
        writer.write(buffer)
        buffer.seek(0)

    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        for p in tmp_paths:
            if os.path.exists(p):
                os.remove(p)

    return send_file(
        buffer,
        mimetype='application/pdf',
        as_attachment=True,
        download_name='merged.pdf',
        headers={
            'X-Files-Merged': str(len(files)),
            'X-Total-Pages':  str(total_pages)
        }
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=True, host='0.0.0.0', port=port)
