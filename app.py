# Application Flask pour uploader un PDF, le traiter avec process_pdf (ocr.py)
# et renvoyer un fichier Excel généré.
import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify, send_file
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader  # pour détecter si le PDF est scanné ou non

# On importe la fonction process_pdf que tu as dans ocr.py (version SANS OCR/Camelot)
from OCR import process_pdf
from extract_with_gemini import process_inputs
# --- Configuration Flask ---
app = Flask(__name__, template_folder='templates')
app.secret_key = "dev_secret_key_change_this"  # change en production

# Dossiers pour stocker les uploads et les résultats
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'result'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# Extensions autorisées et taille max (optionnel)
ALLOWED_EXT = {'pdf', 'png', 'jpg', 'jpeg'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100 MB max upload
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH


def allowed_file(filename):
    """Retourne True si l'extension du fichier est autorisée."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT


def detect_scanned_pdf(filepath):
    """
    Retourne True si le PDF est scanné (image) donc sans texte.
    Retourne False si le PDF contient du texte (natif).
    """
    try:
        reader = PdfReader(filepath)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted
        return len(text.strip()) == 0  # aucun texte trouvé → scanné
    except Exception:
        return True  # en cas d’erreur de lecture, on suppose scanné


# --- Route principale : upload du fichier ---
@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Affiche le formulaire d'upload.
    Si POST et fichier valide -> sauvegarde le fichier et affiche le nom pour extraction.
    """
    filename = None
    is_scanned = False

    if request.method == 'POST':
        # Vérifie que le champ 'file' est présent
        if 'file' not in request.files:
            flash("Aucun fichier reçu (champ 'file' manquant).")
            return redirect(request.url)

        file = request.files['file']

        # Vérifie que l'utilisateur a sélectionné un fichier
        if file.filename == '':
            flash("Aucun fichier choisi.")
            return redirect(request.url)

        # Vérifie l'extension autorisée
        if not allowed_file(file.filename):
            flash("Format non autorisé. Seuls les fichiers PDF sont acceptés.")
            return redirect(request.url)

        # Vérifie la taille du fichier (déjà limitée par MAX_CONTENT_LENGTH, mais check amiable)
        file.seek(0, os.SEEK_END)
        size_mb = file.tell() / (1024 * 1024)
        file.seek(0)
        if size_mb > (MAX_CONTENT_LENGTH / (1024 * 1024)):
            flash(f"Fichier trop grand ({size_mb:.1f} MB). Limite: {MAX_CONTENT_LENGTH/(1024*1024)} MB.")
            return redirect(request.url)

        # Sécurise et rends le nom unique pour éviter collisions
        original_name = secure_filename(file.filename)
        unique_name = f"{uuid.uuid4().hex}_{original_name}"
        save_path = os.path.join(UPLOAD_FOLDER, unique_name)

        # Sauvegarde le fichier uploadé
        file.save(save_path)
        filename = unique_name

        # Détection scanné ou natif
        is_scanned = detect_scanned_pdf(save_path)

        if is_scanned:
            flash(f"Fichier uploadé avec succès (PDF scanné détecté) : {original_name}")
        else:
            flash(f"Fichier uploadé avec succès (PDF natif détecté) : {original_name}")

    # Rend la page d'accueil (template index.html) en envoyant le nom du fichier uploadé si présent
    return render_template('index.html', filename=filename, is_scanned=is_scanned)


# --- Route d'extraction : appelle process_pdf et renvoie l'Excel ---
@app.route('/extract_excel', methods=['POST'])
def extract_excel():
    """
    Récupère le nom du fichier (hidden input from index.html),
    appelle process_pdf(pdf_path) et renvoie le fichier Excel généré.
    """
    filename = request.form.get('filename')
    if not filename:
        flash("Aucun fichier sélectionné pour extraction.")
        return redirect(url_for('index'))

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(pdf_path):
        flash("Fichier introuvable sur le serveur.")
        return redirect(url_for('index'))

    try:
        out_path = process_pdf(
            pdf_path,
            out=os.path.join(RESULT_FOLDER, f"{os.path.splitext(filename)[0]}.xlsx")
        )
    except Exception as e:
        flash(f"Erreur durant le traitement : {e}")
        return redirect(url_for('index'))

    return send_from_directory(RESULT_FOLDER, os.path.basename(out_path), as_attachment=True)


# --- Route pour extraire le texte d'un PDF natif uniquement ---
@app.route('/extract_text_native', methods=['POST'])
def extract_text_native():
    """
    Récupère le nom du fichier (hidden input from index.html),
    appelle pdf_to_text(pdf_path) en mode natif, sauvegarde le texte dans le dossier 'result',
    et propose le téléchargement.
    """
    filename = request.form.get('filename')
    if not filename:
        flash("Aucun fichier sélectionné pour extraction.")
        return redirect(url_for('index'))

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(pdf_path):
        flash("Fichier introuvable sur le serveur.")
        return redirect(url_for('index'))

    try:
        print("Extraction texte natif lancé")
        from OCR import extract_text_native
        text = extract_text_native(pdf_path)
        base_name = os.path.splitext(os.path.basename(filename))[0]
        txt_filename = f"{base_name}_texte_natif.txt"
        txt_path = os.path.join(RESULT_FOLDER, txt_filename)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception as e:
        flash(f"Erreur durant le traitement texte natif : {e}")
        return redirect(url_for('index'))

    return send_from_directory(RESULT_FOLDER, txt_filename, as_attachment=True)


# --- Route pour supprimer un fichier uploadé (AJAX) ---
@app.route('/delete_uploaded_file', methods=['POST'])
def delete_uploaded_file():
    filename = request.json.get('filename')
    if not filename:
        return jsonify({'success': False, 'error': 'Aucun fichier spécifié.'}), 400
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.isfile(file_path):
        try:
            os.remove(file_path)
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 500
    else:
        return jsonify({'success': False, 'error': 'Fichier introuvable.'}), 404


# --- Route d'extraction : appelle process_pdf_gemini et renvoie l'Excel ---
@app.route('/extract_excel_gemini', methods=['POST'])
def extract_excel_gemini():
    """
    Extraction OCR Gemini : traite le PDF ou l'image uploadé(e) avec Gemini et renvoie l'Excel.
    """
    filename = request.form.get('filename')
    if not filename:
        flash("Aucun fichier sélectionné pour extraction.")
        return redirect(url_for('index'))

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(pdf_path):
        flash("Fichier introuvable sur le serveur.")
        return redirect(url_for('index'))

    # Conversion image → PDF si besoin (déjà vu plus haut)
    import mimetypes
    mime, _ = mimetypes.guess_type(pdf_path)
    if mime and mime.startswith("image/"):
        from PIL import Image
        img = Image.open(pdf_path)
        pdf_converted = os.path.splitext(pdf_path)[0] + "_converted.pdf"
        img.convert("RGB").save(pdf_converted, "PDF")
        pdf_path = pdf_converted

    out_xlsx = os.path.join(RESULT_FOLDER, f"{os.path.splitext(filename)[0]}_gemini.xlsx")
    try:
        result_xlsx = process_inputs(pdf_path, out_xlsx=out_xlsx)
    except Exception as e:
        flash(f"Erreur Gemini OCR : {e}")
        return redirect(url_for('index'))

    if not os.path.isfile(result_xlsx):
        flash("Le fichier Excel n'a pas été généré. Vérifiez le traitement Gemini.")
        return redirect(url_for('index'))

    return send_file(out_xlsx, as_attachment=True)


# --- Lancement ---
if __name__ == "__main__":
    app.run(debug=True, port=5001)
