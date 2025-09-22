# create_admin.py
import getpass
import mysql.connector
from werkzeug.security import generate_password_hash

# --- CONFIG DB (adapter si besoin) ---
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',    # ton mot de passe MySQL
    'database': 'ocr_system'
}

def get_db_conn():
    return mysql.connector.connect(**DB_CONFIG)

def create_or_update_admin(email, plain_password):
    hashed = generate_password_hash(plain_password)  # pbkdf2:sha256:... par défaut
    conn = get_db_conn()
    cursor = conn.cursor()
    # vérifie si l'email existe
    cursor.execute("SELECT id FROM admin WHERE email = %s", (email,))
    row = cursor.fetchone()
    if row:
        # update
        cursor.execute("UPDATE admin SET mot_de_passe = %s WHERE email = %s", (hashed, email))
        print(f"Mot de passe mis à jour pour {email}")
    else:
        # insert (ajuste les colonnes selon ta table)
        cursor.execute("INSERT INTO admin (email, mot_de_passe) VALUES (%s, %s)", (email, hashed))
        print(f"Admin créé : {email}")
    conn.commit()
    cursor.close()
    conn.close()


if __name__ == "__main__":
    print("Création / mise à jour d'un admin (hash pbkdf2 via werkzeug)")
    email = input("Email admin: ").strip()
    if not email:
        print("Email requis.")
        exit(1)
    # password en masqué
    pwd = getpass.getpass("Mot de passe : ")
    pwd2 = getpass.getpass("Confirme mot de passe : ")
    if pwd != pwd2:
        print("Les 2 mots de passe ne correspondent pas.")
        exit(1)
    create_or_update_admin(email, pwd)
    print("Terminé.")
