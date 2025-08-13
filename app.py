# Application Flask pour uploader un PDF, le traiter avec process_pdf (ocr.py)
# et renvoyer un fichier Excel généré.
import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify
from werkzeug.utils import secure_filename

# On importe la fonction process_pdf que tu as dans ocr.py (version SANS OCR/Camelot)
from OCR import process_pdf

# --- Configuration Flask ---
app = Flask(__name__, template_folder='templates')
app.secret_key = "dev_secret_key_change_this"  # change en production

# Dossiers pour stocker les uploads et les résultats
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'result'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

# Extensions autorisées et taille max (optionnel)
ALLOWED_EXT = {'pdf'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100 MB max upload
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

def allowed_file(filename):
    """Retourne True si l'extension du fichier est autorisée."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

# --- Route principale : upload du fichier ---
@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Affiche le formulaire d'upload.
    Si POST et fichier valide -> sauvegarde le fichier et affiche le nom pour extraction.
    """
    uploaded_name = None

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
        uploaded_name = unique_name
        flash(f"Fichier uploadé avec succès : {original_name}")

    # Rend la page d'accueil (template index.html) en envoyant le nom du fichier uploadé si présent
    return render_template('index.html', filename=uploaded_name)

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

    # Paramètres pour process_pdf
    pages = "all"
    debug = False
    result_folder = RESULT_FOLDER
    out_name = None  # None -> process_pdf générera un nom basé sur le PDF
    remove_input = True  # supprime le PDF uploadé après traitement (change si tu veux garder)

    try:
        # Appel du module OCR (sans OCR en fait) qui retourne le chemin de l'Excel
        out_path = process_pdf(pdf_path, pages=pages, debug=debug,
                               result_folder=result_folder, out_name=out_name,
                               remove_input=remove_input)
    except Exception as e:
        # Gestion simple d'erreur : flash + redirection à l'index
        flash(f"Erreur durant le traitement : {e}")
        return redirect(url_for('index'))

    # Envoie le fichier Excel en téléchargement
    return send_from_directory(result_folder, os.path.basename(out_path), as_attachment=True)

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

# --- Lancement ---
if __name__ == "__main__":
    app.run(debug=True, port=5001)
