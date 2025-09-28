import mysql.connector
from workflow import process_single_file
import shutil
import os



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
app.secret_key = "dev_secret_key_change_this" 

from admin_auth import admin_bp
app.register_blueprint(admin_bp)

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


import fitz  # PyMuPDF

def detect_scanned_pdf(pdf_path, min_text_len=50, min_text_pages_ratio=0.5):
    doc = fitz.open(pdf_path)
    text_pages = 0
    total_pages = len(doc)  # <--- récupère le nombre de pages AVANT de fermer
    for page in doc:
        txt = page.get_text("text").strip()
        if len(txt) > min_text_len:
            text_pages += 1
    doc.close()
    ratio = text_pages / total_pages  # <--- utilise total_pages ici
    # Si moins de la moitié des pages ont du texte, on considère comme scanné
    return ratio < min_text_pages_ratio



# --- Route principale : upload du fichier ---
@app.route('/', methods=['GET', 'POST'])
def index():
    clients = get_clients()
    filename = None
    client_id = None
    client_name = None
    is_scanned = False
    if request.method == 'POST':
        client_id = request.form.get('client_id')
        for c in clients:
            if str(c[0]) == str(client_id):
                client_name = c[1]
                break
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
            flash(f"Fichier uploadé avec succès (PDF scanné détecté) : {original_name}", "success")
        else:
            flash(f"Fichier uploadé avec succès (PDF natif détecté) : {original_name}", "success")

    # Rend la page d'accueil (template index.html) en envoyant le nom du fichier uploadé si présent
    return render_template('index.html', clients=clients, filename=filename, client_id=client_id, client_name=client_name, is_scanned=is_scanned)


# --- Route d'extraction : appelle process_pdf et renvoie l'Excel ---
@app.route('/extract_excel', methods=['POST'])
def extract_excel():
    filename = request.form.get('filename')
    client_id = request.form.get('client_id')
    client_name = None
    for c in get_clients():
        if str(c[0]) == str(client_id):
            client_name = c[1]
            break

    if not filename:
        flash("Aucun fichier sélectionné pour extraction.")
        return redirect(url_for('index'))

    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(pdf_path):
        flash("Fichier introuvable sur le serveur.")
        return redirect(url_for('index'))

    # --- NOM DU FICHIER EXCEL ---
    if client_id and client_name:
        excel_filename = f"{client_id}-{client_name}.xlsx".replace(" ", "_")
    else:
        excel_filename = f"{os.path.splitext(filename)[0]}.xlsx"
    out_xlsx = os.path.join(RESULT_FOLDER, excel_filename)

    try:
        res = process_pdf(pdf_path, out=out_xlsx)
    except Exception as e:
        app.logger.exception("Erreur durant process_pdf: %s", e)
        flash(f"Erreur durant le traitement : {e}")
        return redirect(url_for('index'))

    result_path = res if isinstance(res, str) and os.path.isfile(res) else (out_xlsx if os.path.isfile(out_xlsx) else None)
    if not result_path:
        flash("Le fichier Excel n'a pas été généré. Vérifiez le traitement.")
        return redirect(url_for('index'))

    # Affiche la page avec les deux boutons
    return render_template(
        'index.html',
        clients=get_clients(),
        excel_ready=True,
        excel_filename=os.path.basename(out_xlsx),
        filename=filename,
        client_id=client_id,
        client_name=client_name,
        is_scanned=False
    )


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
    filename = request.form.get('filename')
    client_id = request.form.get('client_id')
    client_name = None
    for c in get_clients():
        if str(c[0]) == str(client_id):
            client_name = c[1]
            break

    # --- NOM DU FICHIER EXCEL ---
    if client_id and client_name:
        excel_filename = f"{client_id}-{client_name}.xlsx".replace(" ", "_")
    else:
        excel_filename = "result.xlsx"
    out_xlsx = os.path.join(RESULT_FOLDER, excel_filename)

    # --- Chemin du PDF uploadé ---
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.isfile(pdf_path):
        flash("Fichier PDF introuvable.")
        return redirect(url_for('index'))

    # --- Extraction Gemini ---
    try:
        result_xlsx = process_inputs(pdf_path, out_xlsx=out_xlsx, client_id=client_id)
    except Exception as e:
        print(f"Erreur durant l'extraction Gemini : {e}")
        flash(f"Erreur durant l'extraction Gemini : {e}")
        return redirect(url_for('index'))

    print(f"DEBUG: Chemin du fichier Excel généré: {result_xlsx}")
    if not os.path.isfile(result_xlsx):
        print("DEBUG: Fichier Excel non trouvé !")
        flash("Le fichier Excel n'a pas été généré. Vérifiez le traitement.")
        return redirect(url_for('index'))

    # --- Affiche la page avec les deux boutons ---
    return render_template(
        "index.html",
        clients=get_clients(),
        excel_ready=True,
        excel_filename=excel_filename,
        filename=filename,
        client_id=client_id,
        client_name=client_name,
        is_scanned=True
    )


@app.route('/download_excel', methods=['POST'])
def download_excel():
    filename = request.form.get('filename')
    path = os.path.join(RESULT_FOLDER, filename)
    return send_file(path, as_attachment=True)

@app.route('/save_excel_to_db', methods=['POST'])
def save_excel_to_db():
    filename = request.form.get('filename')
    excel_path = os.path.join(RESULT_FOLDER, filename)
    proceed_path = os.path.join("proceed", filename)
    try:
        process_single_file(excel_path)
        shutil.copy2(excel_path, proceed_path)
        # Essaye de supprimer plusieurs fois si le fichier est verrouillé
        deleted = False
        for _ in range(3):
            try:
                os.remove(excel_path)
                deleted = True
                break
            except PermissionError:
                import time
                time.sleep(1)
        else:
            flash("Fichier sauvegardé dans la BDD et déplacé vers proceed !", "success")
    except Exception as e:
        flash(f"Erreur lors de la sauvegarde : {e}", "danger")
    return redirect(url_for('index'))


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
    app.run(debug=True, port=5003)
