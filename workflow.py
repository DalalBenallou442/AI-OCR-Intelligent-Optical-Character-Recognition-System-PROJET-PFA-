# -*- coding: utf-8 -*-
"""
workflow.py - Sauvegarde des bilans Excel (Actif, Passif, CPC) dans MySQL
"""

import os
import argparse
import logging
import pandas as pd
import mysql.connector as mysql
from flask import flash, redirect, url_for, request

# Config logs
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# -----------------------
# Connexion MySQL
# -----------------------
def get_db_connection():
    return mysql.connect(
        host="localhost",
        user="root",
        password="",  # adapte si tu as un mot de passe
        database="ocr_system"
    )

# -----------------------
# Nettoyage DataFrame
# -----------------------
def clean_dataframe(df):
    df.columns = (
        df.columns.str.strip()
        .str.replace("'", "_", regex=True)  # Remplace apostrophe par underscore
        .str.replace(" ", "_", regex=True)
        .str.replace("[^0-9a-zA-Z_]", "", regex=True)
        .str.lower()
    )
    return df

# -----------------------
# Insertion SQL
# -----------------------
def insert_rows_into_db(conn, table_name, df):
    print(f"DEBUG: insert_rows_into_db table={table_name}, lignes={len(df)}")
    if df.empty:
        logging.warning(f"Aucune donnée à insérer dans {table_name}")
        return 0

    cursor = conn.cursor()

    # Préparer colonnes & placeholders
    cols = ", ".join([f"`{c}`" for c in df.columns])
    placeholders = ", ".join(["%s"] * len(df.columns))
    insert_query = f"INSERT INTO {table_name} ({cols}) VALUES ({placeholders})"

    inserted = 0
    for row in df.itertuples(index=False, name=None):
        try:
            # Remplacer NaN par None
            cleaned_row = [None if (pd.isna(x) or str(x).lower() == "nan") else x for x in row]
            cursor.execute(insert_query, cleaned_row)
            conn.commit()
            inserted += 1
        except Exception as e:
            logging.error(f"Erreur insertion dans {table_name}: {e}")
            logging.error(f"Ligne: {row}")
    return inserted

# -----------------------
# Traitement d'un fichier
# -----------------------
def process_single_file(file_path):
    print(f"DEBUG: process_single_file appelé avec {file_path}")
    logging.info(f"📂 Traitement du fichier: {file_path}")

    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        logging.error(f"Impossible de lire {file_path}: {e}")
        return

    conn = get_db_connection()

    # Extraire l'id_client du nom de fichier (ex: 100-Zine_Cereales.xlsx → 100)
    base = os.path.basename(file_path)
    id_client = base.split("-", 1)[0] if "-" in base else None

    # Mapping: feuille Excel -> table MySQL
    mapping = {
        "Bilan_Actif": "bilan_actif",
        "Bilan_Passif": "bilan_passif",
        "Bilan_CPC": "cpc"
    }

    table_columns_mapping = {
        "bilan_actif": [
            "id_client",
            "type_tableau",
            "parent_sous_categorie",
            "sous_categorie",
            "rubrique",
            "montant_brut",
            "amortissements_provisions",
            "net_exercice",
            "net_exercice_prec",
            "commentaires",
            "matched_from_page",
            "matched_name_raw",
            "source"
        ],
        "bilan_passif": [
            "id_client",
            "type_tableau",
            "parent_sous_categorie",
            "sous_categorie",
            "rubrique",
            "net_exercice",
            "net_exercice_prec",
            "commentaires",
            "matched_from_page",
            "matched_name_raw",
            "source"
        ],
        "cpc": [
            "id_client",
            "type_tableau",
            "parent_sous_categorie",
            "sous_categorie",
            "rubrique",
            "propres_a_l_exercice",
            "concernant_les_exercices_prec",
            "totaux_de_l_exercice",
            "exercice_prec",
            "commentaires",
            "_matched_from_page",
            "_matched_name_raw",
            "_source"
        ]
    }

    for sheet, table in mapping.items():
        if sheet in xls.sheet_names:
            logging.info(f"➡ Feuille {sheet} détectée → table {table}")
            try:
                df = pd.read_excel(file_path, sheet_name=sheet)
                df = clean_dataframe(df)
                df["id_client"] = id_client

                # Utilise la bonne liste de colonnes
                table_columns = table_columns_mapping[table]
                df = df.reindex(columns=table_columns)

                logging.info("Colonnes DataFrame: %s", df.columns.tolist())
                cnt = insert_rows_into_db(conn, table, df)
                logging.info(f"✅ {cnt} lignes insérées dans {table}")
            except Exception as e:
                logging.error(f"ERREUR traitement {sheet}: {e}")
        else:
            logging.warning(f"⚠ Feuille {sheet} absente du fichier.")

    conn.close()

# -----------------------
# Routes Flask
# -----------------------


# -----------------------
# Main
# -----------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Workflow import Excel vers MySQL")
    parser.add_argument("--file", help="Chemin du fichier Excel à traiter", required=True)
    args = parser.parse_args()

    if not os.path.exists(args.file):
        logging.error(f"Fichier introuvable: {args.file}")
    else:
        process_single_file(args.file)
