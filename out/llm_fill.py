#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
llm_fill_exact.py
Mapping strict des rubriques et sous-catégories
sans utiliser RapidFuzz pour éviter les erreurs de correspondance
"""

import pandas as pd
import re

TARGET_COLUMNS = [
    "Type_Tableau", "Sous_Catégorie", "Rubrique",
    "Montant_Brut", "Amortissements_Provisions",
    "Net_Exercice", "Net_Exercice_Prec", "Commentaires"
]

DESIRED_ROWS = [
    ("IMMOBILISATIONS EN NON VALEUR","Frais préliminaires"),
    ("IMMOBILISATIONS EN NON VALEUR","Charges à répartir sur plusieurs exercices"),
    ("IMMOBILISATIONS INCORPORELLES","Immobilisations en recherche et développement"),
    ("IMMOBILISATIONS INCORPORELLES","Brevets, marques, droits et valeurs similaires"),
    ("IMMOBILISATIONS CORPORELLES","Terrains"),
    ("IMMOBILISATIONS CORPORELLES","Constructions"),
    ("IMMOBILISATIONS CORPORELLES","Installations techniques, matériel et outillage"),
    ("IMMOBILISATIONS CORPORELLES","Matériel de transport"),
    ("IMMOBILISATIONS CORPORELLES","Mobilier, matériel de bureau et aménagements divers"),
    ("IMMOBILISATIONS CORPORELLES","Autres immobilisations corporelles"),
    ("IMMOBILISATIONS CORPORELLES","Immobilisations corporelles en cours"),
    ("IMMOBILISATIONS FINANCIÈRES","Prêts immobilisés"),
    ("IMMOBILISATIONS FINANCIÈRES","Autres créances financières"),
    ("IMMOBILISATIONS FINANCIÈRES","Titres de participation"),
    ("IMMOBILISATIONS FINANCIÈRES","Autres titres immobilisés"),
    ("ÉCARTS DE CONVERSION – ACTIF","Diminution des dettes immobilisations"),
    ("ÉCARTS DE CONVERSION – ACTIF","Augmentation des créances immobilisations"),
    ("STOCKS","Marchandises"),
    ("STOCKS","Matières et fournitures consommables"),
    ("STOCKS","Produits intermédiaires et résiduels"),
    ("STOCKS","Produits finis"),
    ("CRÉANCES ACTIF CIRCULANT","Fournisseurs débiteurs, avances et acomptes"),
    ("CRÉANCES ACTIF CIRCULANT","Clients et comptes rattachés"),
    ("CRÉANCES ACTIF CIRCULANT","Personnel"),
    ("CRÉANCES ACTIF CIRCULANT","État"),
    ("CRÉANCES ACTIF CIRCULANT","Comptes d’associés"),
    ("CRÉANCES ACTIF CIRCULANT","Autres débiteurs"),
    ("CRÉANCES ACTIF CIRCULANT","Comptes de régularisation actif"),
    ("TITRES ET VALEURS DE PLACEMENT","Titres et valeurs de placement"),
    ("TRÉSORERIE ACTIF","Chèques et valeurs à encaisser"),
    ("TRÉSORERIE ACTIF","Banques, Trésor et CCP"),
    ("TRÉSORERIE ACTIF","Caisse, régies d’avances et accréditifs"),
]

def normalize(s):
    if s is None: return ""
    s = str(s).strip().lower()
    s = re.sub(r'[éèêàùç]', lambda x: {'é':'e','è':'e','ê':'e','à':'a','ù':'u','ç':'c'}[x.group()], s)
    s = s.replace("’","'").replace("´","'")
    s = re.sub(r'\s+', ' ', s)
    return s

def map_rows(input_csv, output_csv):
    df = pd.read_csv(input_csv, sep=';', encoding='utf-8-sig').fillna('')
    df['__norm'] = (df['Sous_Catégorie'].fillna('') + '||' + df['Rubrique'].fillna('')).map(normalize)

    rows_out = []
    for s, r in DESIRED_ROWS:
        key = normalize(f"{s}||{r}")
        matched = df[df['__norm'] == key]
        if matched.empty:
            # fallback substring search strict
            matched = df[df['__norm'].str.contains(normalize(r)) | df['__norm'].str.contains(normalize(s))]
        row_data = matched.iloc[0] if not matched.empty else None
        rows_out.append({
            "Type_Tableau": "Bilan_Actif",
            "Sous_Catégorie": s,
            "Rubrique": r,
            "Montant_Brut": row_data['Montant_Brut'] if row_data is not None else "",
            "Amortissements_Provisions": row_data['Amortissements_Provisions'] if row_data is not None else "",
            "Net_Exercice": row_data['Net_Exercice'] if row_data is not None else "",
            "Net_Exercice_Prec": row_data['Net_Exercice_Prec'] if row_data is not None else "",
            "Commentaires": row_data['Commentaires'] if row_data is not None else ""
        })

    df_final = pd.DataFrame(rows_out, columns=TARGET_COLUMNS)
    df_final.to_csv(output_csv, index=False, sep=';', encoding='utf-8-sig', quoting=1)
    print(f"[+] CSV final sauvegardé : {output_csv}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Map exact DESIRED_ROWS from OCR CSV")
    parser.add_argument("--csv_in", required=True, help="CSV OCR pré-extrait")
    parser.add_argument("--csv_out", required=True, help="CSV final structuré")
    args = parser.parse_args()

    map_rows(args.csv_in, args.csv_out)
