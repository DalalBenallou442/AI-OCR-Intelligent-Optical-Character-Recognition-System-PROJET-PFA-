import mysql.connector

def get_clients():
    """Récupère la liste des clients depuis la table CLIENTS"""
    try:
        conn = mysql.connector.connect(
            host="localhost",      # ou l’IP de ton serveur MySQL
            user="root",           # ton user MySQL
            password="", # ton mot de passe MySQL
            database="ocr_system"
        )
        cursor = conn.cursor()
        cursor.execute("SELECT ID_CLIENT, NOM_CLIENT FROM CLIENT ORDER BY ID_CLIENT")
        clients = cursor.fetchall()  
        cursor.close()
        conn.close()
        return clients
    except Exception as e:
        print("Erreur MySQL:", e)
        return []

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


def detect_scanned_pdf(filepath, min_alpha_chars=30):
    """
    Retourne True si le PDF semble scanné (image) :
    - extrait le texte (PyPDF2) et compte les caractères alphabétiques ;
    - si peu de texte -> vérifie la présence d'images via fitz (page.get_images()).
    """
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(filepath)
        text = ""
        for page in reader.pages:
            txt = page.extract_text()
            if txt:
                text += txt + "\n"

        alpha_chars = sum(1 for c in text if c.isalpha())
        if alpha_chars >= min_alpha_chars:
            # beaucoup de texte => PDF natif
            return False

        # si peu de texte on considère scanné, mais on vérifie la présence d'images
        try:
            import fitz
            doc = fitz.open(filepath)
            img_count = 0
            for page in doc:
                imgs = page.get_images(full=True)
                if imgs:
                    img_count += len(imgs)
            doc.close()
            # s'il y a des images => scanné
            return img_count > 0 or alpha_chars < min_alpha_chars
        except Exception:
            # si fitz indisponible, on parie sur le peu de texte et on retourne True
            return True

    except Exception:
        # en cas d'erreur on suppose scanné (comportement sûr)
        return True



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
        client_id = request.form.get('client_id')

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
    clients = get_clients()
    return render_template('index.html', filename=filename, is_scanned=is_scanned, clients=clients)


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
    Accepts optional form field 'force_scanned' to force "scanned" processing.
    """
    filename = request.form.get('filename')
    if not filename:
        flash("Aucun fichier sélectionné pour extraction.")
        return redirect(url_for('index'))

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(pdf_path):
        flash("Fichier introuvable sur le serveur.")
        return redirect(url_for('index'))

    # --- Lecture du paramètre facultatif "forcer scanné" ---
    force_val = request.form.get("force_scanned")
    force = False
    if force_val is not None:
        # accepte "1", "on", "true", "yes"
        if str(force_val).lower() in ("1", "on", "true", "yes"):
            force = True

    # Détection (sauf si forcée)
    if not force:
        is_scanned_local = detect_scanned_pdf(pdf_path)
    else:
        is_scanned_local = True
        print("DEBUG: Forcé par l'utilisateur -> traitement comme PDF scanné")

    # Si c'est une image unique upload (ex: image/png) converti en PDF
    import mimetypes
    mime, _ = mimetypes.guess_type(pdf_path)
    if mime and mime.startswith("image/") and not pdf_path.lower().endswith(".pdf"):
        from PIL import Image
        img = Image.open(pdf_path)
        pdf_converted = os.path.splitext(pdf_path)[0] + "_converted.pdf"
        img.convert("RGB").save(pdf_converted, "PDF")
        pdf_path = pdf_converted
        is_scanned_local = True

    out_xlsx = os.path.join(RESULT_FOLDER, f"{os.path.splitext(filename)[0]}_gemini.xlsx")
    try:
        if is_scanned_local:
            result_xlsx = process_inputs(pdf_path, out_xlsx=out_xlsx)
        else:
            result_xlsx = process_pdf(pdf_path, out_xlsx=out_xlsx)

        # AJOUT DEBUG
        print(f"DEBUG Flask: result_xlsx = {result_xlsx}")
        print(f"DEBUG Flask: result_xlsx exists? {os.path.exists(result_xlsx) if result_xlsx else 'result_xlsx is None'}")

        if not result_xlsx:
            flash("Aucun résultat extrait par Gemini OCR.")
            return redirect(url_for('index'))
    except Exception as e:
        print(f"DEBUG Flask: Exception = {e}")  # AJOUT
        flash(f"Erreur Gemini OCR : {e}")
        return redirect(url_for('index'))

    if not os.path.isfile(result_xlsx):
        print(f"DEBUG Flask: Le fichier {result_xlsx} n'existe pas")  # AJOUT
        flash("Le fichier Excel n'a pas été généré. Vérifiez le traitement.")
        return redirect(url_for('index'))

    print(f"DEBUG Flask: About to send file: {result_xlsx}")  # AJOUT
    return send_file(result_xlsx, as_attachment=True)




# --- Lancement ---
if __name__ == "__main__":
    # Test de connexion MySQL
    print("Test connexion MySQL...")
    clients = get_clients()
    if clients:
        print(f"Connexion OK, {len(clients)} clients trouvés.")
    else:
        print("Connexion échouée ou aucun client trouvé.")
    # Tu peux ensuite lancer Flask normalement
    app.run(debug=True, port=5001)
