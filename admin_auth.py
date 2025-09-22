# admin_auth.py
from flask import Blueprint, request, render_template, redirect, url_for, flash, session
import mysql.connector
from werkzeug.security import check_password_hash, generate_password_hash

admin_bp = Blueprint('admin', __name__, template_folder='templates')

# --- CONFIG DB (adapte si nécessaire) ---
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',       # ton mot de passe MySQL si nécessaire
    'database': 'ocr_system'
}

def get_db_conn():
    return mysql.connector.connect(**DB_CONFIG)

# ROUTE: Login (email + mot de passe)
@admin_bp.route('/admin/login', methods=['GET', 'POST'])
def login():
    # si déjà connecté -> redirige vers dashboard
    if session.get('admin_authenticated'):
        return redirect(url_for('admin.clients_list'))

    if request.method == 'POST':
        email = (request.form.get('email') or '').strip().lower()
        password = request.form.get('password') or ''

        if not email or not password:
            flash("Remplis l'email et le mot de passe.", "warning")
            return render_template('admin_login.html')

        try:
            conn = get_db_conn()
            cursor = conn.cursor(dictionary=True)
            # ATTENTION : adapte le nom de la colonne de mot de passe selon ta table.
            # Tu m'as montré : colonnes id, email, mot_de_passe
            cursor.execute("SELECT id, email, mot_de_passe FROM admin WHERE email = %s", (email,))
            row = cursor.fetchone()
            cursor.close()
            conn.close()
        except Exception:
            flash("Erreur serveur (BD).", "danger")
            return render_template('admin_login.html')

        if not row:
            flash("Email introuvable.", "danger")
            return render_template('admin_login.html')

        # mot_de_passe peut être un hash ou un mot de passe en clair.
        stored = row.get('mot_de_passe') or ''
        password_ok = False

        # 1) si stored semble déjà être un hash werkzeug (pbkdf2:, scrypt:, argon2:)
        if isinstance(stored, str) and (stored.startswith("pbkdf2:") or stored.startswith("scrypt:") or stored.startswith("argon2:")):
            password_ok = check_password_hash(stored, password)
        else:
            # stored n'est pas un hash : on compare en clair
            if password == stored:
                password_ok = True
                # -> upgrade : remplace le mot de passe en clair par un hash sécurisé dans la DB
                try:
                    new_hash = generate_password_hash(password)  # pbkdf2:sha256 par défaut
                    conn = get_db_conn()
                    cur = conn.cursor()
                    cur.execute("UPDATE admin SET mot_de_passe = %s WHERE id = %s", (new_hash, row['id']))
                    conn.commit()
                    cur.close()
                    conn.close()
                except Exception:
                    # si update échoue, on ignore l'upgrade mais l'utilisateur est connecté
                    pass

        if not password_ok:
            flash("Mot de passe incorrect.", "danger")
            return render_template('admin_login.html')

        # ----- pour le moment on authentifie directement la session -----
        session['admin_authenticated'] = True
        session['admin_id'] = row['id']
        session['admin_email'] = row['email']
        flash("Connexion réussie.", "success")
        return redirect(url_for('admin.clients_list'))  # <-- ici !

    return render_template('admin_login.html')


# ROUTE: Admin Dashboard (protégée)
@admin_bp.route('/admin')
def admin_dashboard():
    if not session.get('admin_authenticated'):
        return redirect(url_for('admin.login'))
    return render_template('admin_dashboard.html')


# ROUTE: Logout
@admin_bp.route('/admin/logout')
def logout():
    session.pop('admin_authenticated', None)
    session.pop('admin_id', None)
    session.pop('admin_email', None)
    flash("Déconnecté.", "info")
    return redirect(url_for('index'))


@admin_bp.route('/admin/clients', methods=['GET', 'POST'])
def clients_list():
    if request.method == 'POST':
        id_client = request.args.get('edit_id')
        nom_client = request.form.get('nom_client')
        if id_client and nom_client:
            try:
                conn = get_db_conn()
                cur = conn.cursor()
                cur.execute("UPDATE CLIENT SET NOM_CLIENT=%s WHERE ID_CLIENT=%s", (nom_client, id_client))
                conn.commit()
                cur.close()
                flash("Client mis à jour.", "success")
            except Exception as e:
                flash(f"Erreur update : {e}", "danger")
            finally:
                try: conn.close()
                except: pass
        return redirect(url_for('admin.clients_list'))
    id_q = (request.args.get('id_client') or '').strip()
    name_q = (request.args.get('nom_client') or '').strip()
    conn = None
    try:
        conn = get_db_conn()
        cur = conn.cursor(dictionary=True)
        if id_q or name_q:
            sql = "SELECT ID_CLIENT, NOM_CLIENT FROM CLIENT WHERE 1=1"
            params = []
            if id_q:
                sql += " AND ID_CLIENT = %s"
                params.append(id_q)
            if name_q:
                sql += " AND NOM_CLIENT LIKE %s"
                params.append(f"%{name_q}%")
            sql += " ORDER BY ID_CLIENT"
            cur.execute(sql, tuple(params))
        else:
            cur.execute("SELECT ID_CLIENT, NOM_CLIENT FROM CLIENT ORDER BY ID_CLIENT")
        rows = cur.fetchall()
        cur.close()
    except Exception as e:
        flash(f"Erreur BD: {e}", "danger")
        rows = []
    finally:
        if conn:
            conn.close()
    return render_template('clients_list_search.html', clients=rows, id_query=id_q, name_query=name_q)

@admin_bp.route('/admin/clients/add', methods=['POST'])
def client_add():
    id_client = (request.form.get('id_client') or '').strip()
    nom_client = (request.form.get('nom_client') or '').strip()
    if not id_client or not nom_client:
        flash("id_client et nom_client sont requis.", "warning")
        return redirect(url_for('admin.clients_list'))
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute("INSERT INTO CLIENT (ID_CLIENT, NOM_CLIENT) VALUES (%s, %s)", (id_client, nom_client))
        conn.commit()
        cur.close()
        flash("Client ajouté.", "success")
    except Exception as e:
        flash(f"Erreur ajout : {e}", "danger")
    finally:
        try: conn.close()
        except: pass
    return redirect(url_for('admin.clients_list'))

@admin_bp.route('/admin/clients/<id_client>/bilans', methods=['GET'])
def client_bilans(id_client):
    try:
        conn = get_db_conn()
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT ID_CLIENT, NOM_CLIENT FROM CLIENT WHERE ID_CLIENT=%s", (id_client,))
        client = cur.fetchone()
        if not client:
            cur.close()
            flash("Client introuvable.", "warning")
            return redirect(url_for('admin.clients_list'))

        # récupère les bilans (lecture seule)
        cur.execute("SELECT * FROM bilan_actif WHERE id_client=%s ORDER BY matched_from_page, rubrique", (id_client,))
        actif = cur.fetchall()
        cur.execute("SELECT * FROM bilan_passif WHERE id_client=%s ORDER BY rubrique", (id_client,))
        passif = cur.fetchall()
        cur.execute("SELECT * FROM cpc WHERE id_client=%s ORDER BY rubrique", (id_client,))
        cpc = cur.fetchall()
        cur.close()
    except Exception as e:
        flash(f"Erreur BD: {e}", "danger")
        return redirect(url_for('admin.clients_list'))
    finally:
        try: conn.close()
        except: pass

    return render_template('client_bilans_readonly.html', client=client, actif=actif, passif=passif, cpc=cpc)

@admin_bp.route('/admin/clients/edit/<id_client>', methods=['GET', 'POST'])
def client_edit(id_client):
    if request.method == 'POST':
        nom_client = (request.form.get('nom_client') or '').strip()
        if not nom_client:
            flash("nom_client requis.", "warning")
            return redirect(url_for('admin.client_edit', id_client=id_client))
        try:
            conn = get_db_conn()
            cur = conn.cursor()
            cur.execute("UPDATE CLIENT SET NOM_CLIENT=%s WHERE ID_CLIENT=%s", (nom_client, id_client))
            conn.commit()
            cur.close()
            flash("Client mis à jour.", "success")
            return redirect(url_for('admin.clients_list'))
        except Exception as e:
            flash(f"Erreur update : {e}", "danger")
            return redirect(url_for('admin.client_edit', id_client=id_client))
        finally:
            try: conn.close()
            except: pass
    # GET : affiche le formulaire
    try:
        conn = get_db_conn()
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT ID_CLIENT, NOM_CLIENT FROM CLIENT WHERE ID_CLIENT=%s", (id_client,))
        client = cur.fetchone()
        cur.close()
        if not client:
            flash("Client introuvable.", "warning")
            return redirect(url_for('admin.clients_list'))
    except Exception as e:
        flash(f"Erreur BD: {e}", "danger")
        return redirect(url_for('admin.clients_list'))
    finally:
        try: conn.close()
        except: pass
    return render_template('client_form.html', client=client)

@admin_bp.route('/admin/clients/delete/<id_client>', methods=['POST'])
def client_delete(id_client):
    try:
        conn = get_db_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM bilan_actif WHERE id_client=%s", (id_client,))
        cur.execute("DELETE FROM bilan_passif WHERE id_client=%s", (id_client,))
        cur.execute("DELETE FROM cpc WHERE id_client=%s", (id_client,))
        cur.execute("DELETE FROM CLIENT WHERE ID_CLIENT=%s", (id_client,))
        conn.commit()
        cur.close()
        flash("Client et données associées supprimés.", "success")
    except Exception as e:
        flash(f"Erreur suppression: {e}", "danger")
    finally:
        try: conn.close()
        except: pass
    return redirect(url_for('admin.clients_list'))
