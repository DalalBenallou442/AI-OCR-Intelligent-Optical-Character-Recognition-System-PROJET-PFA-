# extract_with_gemini_full_updated.py
# Script complet — mis à jour : normalisation intelligente des "Sous_Categorie"
import os
import re
import json
import base64
import difflib
import fitz  # PyMuPDF
import pandas as pd
from dotenv import load_dotenv

# Gemini client
import google.generativeai as genai

# ----------------- Config / DEBUG -----------------
load_dotenv()
API_KEYS = [
    os.getenv("GEMINI_API_KEY1"),
    os.getenv("GEMINI_API_KEY2"),
    os.getenv("GEMINI_API_KEY3"),
]
API_KEYS = [k for k in API_KEYS if k]
if not API_KEYS:
    raise RuntimeError("⚠️ Mets au moins une clé dans .env : GEMINI_API_KEY1, GEMINI_API_KEY2, ...")

current_key_index = 0
DEBUG = True  # passe à False en prod

def init_genai():
    global current_key_index
    genai.configure(api_key=API_KEYS[current_key_index])
    if DEBUG:
        print(f"🔑 Utilisation clé API index: {current_key_index}")

def switch_key():
    global current_key_index
    current_key_index = (current_key_index + 1) % len(API_KEYS)
    if DEBUG:
        print(f"⚠️ Bascule vers la clé {current_key_index} (index {current_key_index})")
    init_genai()

init_genai()

# ----------------- Prompt principal -----------------
PROMPT_BILAN_TABLEAU = """
Tu es un assistant OCR intelligent. Tu reçois en entrée une ou plusieurs images (ou un PDF scanné converti en images)
contenant un tableau scanné de bilan comptable en français.

Ignore tout texte avant le tableau.
Renvoie **toujours** une liste JSON (liste de dictionnaires). Pour chaque ligne fournis ces champs EXACTS :
Type_Tableau, Sous_Categorie, Rubrique, Montant_Brut, Amortissements_Provisions, Net_Exercice, Net_Exercice_Prec, Commentaires

Les montants doivent être au format 123456.78 (séparateur point). Si illisible, renvoie "" (chaine vide).

Si tu ne peux pas lire un montant, renvoie "" (vide). Ne renvoie aucun texte hors du JSON.
"""

# ----------------- Appel Gemini -----------------
def call_gemini(content, model_name="gemini-1.5-flash"):
    init_genai()
    try:
        model = genai.GenerativeModel(model_name)
        resp = model.generate_content(content)
        text = getattr(resp, "text", None) or getattr(resp, "output_text", None) or str(resp)
        if DEBUG:
            print("=== Réponse brute Gemini (début) ===")
            print(text[:3000])
            print("=== fin ===")
        return text
    except Exception as e:
        err = str(e).lower()
        if "429" in err or "quota" in err or "rate limit" in err:
            print("🚨 Limite atteinte → changement de clé")
            switch_key()
            return call_gemini(content, model_name)
        raise

# ----------------- Extraction images PDF -----------------
def extract_images_from_pdf(pdf_path, zoom=2):
    doc = fitz.open(pdf_path)
    img_bytes_list = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img_bytes_list.append(pix.tobytes("png"))
    doc.close()
    return img_bytes_list

# ----------------- Safe JSON extract -----------------
def safe_json_extract(text):
    if not text:
        return []
    match = re.search(r"\[.*\]", text, re.S)
    if not match:
        if DEBUG:
            print("⚠️ Aucun JSON détecté dans la réponse Gemini (preview):", text[:1000])
        return []
    raw = match.group(0)
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except Exception as e:
        if DEBUG:
            print("⚠️ JSON non valide après nettoyage:", e)
            print("Réponse brute Gemini (1000 chars):", text[:1000])
        return []

# ----------------- Parsers textuels (fallback) -----------------
_re_number = re.compile(r'[-]?\d{1,3}(?:[ \d]{0,}\d)?[.,]\d{2}')

def normalize_number_str(s):
    if s is None:
        return ""
    s = str(s).strip()
    if s in ["", "-", "—", "–", "--"]:
        return ""
    s = s.replace('\u00A0', ' ')
    m = _re_number.search(s)
    if not m:
        s2 = re.sub(r'[^\d\-,\.\(\)]', '', s)
        if not s2:
            return ""
        s = s2
    else:
        s = m.group(0)
    s = s.replace(' ', '').replace(',', '.')
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    if s.count('.') > 1:
        parts = s.split('.')
        s = ''.join(parts[:-1]) + '.' + parts[-1]
    if re.match(r'^-?\d+(\.\d+)?$', s):
        try:
            return "{:.2f}".format(float(s))
        except:
            return ""
    return ""

def parse_markdown_table(text):
    rows = []
    lines = [ln.rstrip() for ln in text.splitlines()]
    table_lines = []
    for ln in lines:
        if '|' in ln:
            if re.match(r'^\s*\|?\s*-+\s*\|', ln):
                continue
            if re.match(r'^\s*\|?\s*$', ln):
                continue
            table_lines.append(ln)
    if not table_lines:
        return rows
    for ln in table_lines:
        parts = [p.strip() for p in re.split(r'\|', ln)]
        if parts and parts[0] == "":
            parts = parts[1:]
        if parts and parts[-1] == "":
            parts = parts[:-1]
        if len(parts) >= 2:
            rub = parts[0].lstrip('*').strip()
            brut = normalize_number_str(parts[1]) if len(parts) > 1 else ""
            amort = normalize_number_str(parts[2]) if len(parts) > 2 else ""
            net = normalize_number_str(parts[3]) if len(parts) > 3 else ""
            net_prev = normalize_number_str(parts[4]) if len(parts) > 4 else ""
            rows.append({
                "Type_Tableau": "",
                "Sous_Categorie": "",
                "Rubrique": rub,
                "Montant_Brut": brut,
                "Amortissements_Provisions": amort,
                "Net_Exercice": net,
                "Net_Exercice_Prec": net_prev,
                "Commentaires": ""
            })
    return rows

def parse_line_blocks(text):
    rows = []
    lines = [ln for ln in text.splitlines()]
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        if ln.startswith('*'):
            rub = ln.lstrip('*').strip()
            nums = []
            j = i+1
            while j < len(lines) and len(nums) < 4:
                cand = lines[j].strip()
                if cand == "":
                    nums.append("")
                else:
                    if re.search(r'\d', cand):
                        found = _re_number.findall(cand)
                        if found:
                            for nf in found:
                                if len(nums) < 4:
                                    nums.append(normalize_number_str(nf))
                        else:
                            nums.append(normalize_number_str(cand))
                    else:
                        if cand in ['-', '—', '–', '--']:
                            nums.append("")
                        else:
                            break
                j += 1
            while len(nums) < 4:
                nums.append("")
            rows.append({
                "Type_Tableau": "",
                "Sous_Categorie": "",
                "Rubrique": rub,
                "Montant_Brut": nums[0] or "",
                "Amortissements_Provisions": nums[1] or "",
                "Net_Exercice": nums[2] or "",
                "Net_Exercice_Prec": nums[3] or "",
                "Commentaires": ""
            })
            i = j
            continue
        if ln and not re.search(r'\d', ln):
            if i+1 < len(lines):
                next_ln = lines[i+1].strip()
                found = _re_number.findall(next_ln)
                if len(found) >= 3:
                    brut = normalize_number_str(found[0]) if len(found) > 0 else ""
                    amort = normalize_number_str(found[1]) if len(found) > 1 else ""
                    net = normalize_number_str(found[2]) if len(found) > 2 else ""
                    net_prev = normalize_number_str(found[3]) if len(found) > 3 else ""
                    rows.append({
                        "Type_Tableau": "",
                        "Sous_Categorie": ln,
                        "Rubrique": "TOTAL-SOUS-CAT",
                        "Montant_Brut": brut,
                        "Amortissements_Provisions": amort,
                        "Net_Exercice": net,
                        "Net_Exercice_Prec": net_prev,
                        "Commentaires": ""
                    })
                    i += 2
                    continue
        i += 1
    return rows

def parse_gemini_text_to_rows(text):
    rows = []
    md_rows = parse_markdown_table(text)
    if md_rows:
        rows.extend(md_rows)
    block_rows = parse_line_blocks(text)
    if block_rows:
        existing = set((r.get("Rubrique",""), r.get("Sous_Categorie","")) for r in rows)
        for br in block_rows:
            key = (br.get("Rubrique",""), br.get("Sous_Categorie",""))
            if key not in existing:
                rows.append(br)
    if not rows:
        lines = [ln.strip() for ln in text.splitlines() if ln.strip() != ""]
        for idx, ln in enumerate(lines[:-1]):
            if not re.search(r'\d', ln) and idx+1 < len(lines):
                found = _re_number.findall(lines[idx+1])
                if found:
                    brut = normalize_number_str(found[0]) if len(found) > 0 else ""
                    amort = normalize_number_str(found[1]) if len(found) > 1 else ""
                    net = normalize_number_str(found[2]) if len(found) > 2 else ""
                    net_prev = normalize_number_str(found[3]) if len(found) > 3 else ""
                    rows.append({
                        "Type_Tableau": "",
                        "Sous_Categorie": "",
                        "Rubrique": ln.lstrip('*').strip(),
                        "Montant_Brut": brut,
                        "Amortissements_Provisions": amort,
                        "Net_Exercice": net,
                        "Net_Exercice_Prec": net_prev,
                        "Commentaires": ""
                    })
    return rows

# ----------------- Liste des rubriques attendues -----------------
ALL_RUBRIQUES = {
   "Bilan_Actif": {
        "IMMOBILISATIONS EN NON VALEUR (A)": [
            "Frais préliminaires",
            "Charges à répartir sur plusieurs exercices",
            "Primes de remboursement des obligations"
        ],
        "IMMOBILISATIONS INCORPORELLES (B)": [
            "Immobilisations en recherche et développement",
            "Brevets, marques, droits et valeurs similaires",
            "Fonds commercial",
            "Autres immobilisations incorporelles"
        ],
        "IMMOBILISATIONS CORPORELLES (C)": [
            "Terrains",
            "Constructions",
            "Installations techniques, matériel et outillage",
            "Matériel de transport",
            "Mobilier, matériel de bureau et aménagements divers",
            "Autres immobilisations corporelles",
            "Immobilisations corporelles en cours"
        ],
        "IMMOBILISATIONS FINANCIÈRES (D)": [
            "Prêts immobilisés",
            "Autres créances financières",
            "Titres de participation",
            "Autres titres immobilisés"
        ],
        "ÉCARTS DE CONVERSION – ACTIF (E)": [
            "Diminution des créances immobilisées",
            "Augmentation des dettes de financement"
        ],
        "STOCKS (F)": [
            "Marchandises",
            "Matières et fournitures consommables",
            "Produits en cours",
            "Produits intermédiaires et produits résiduels",
            "Produits finis"
        ],
        "CRÉANCES DE L'ACTIF CIRCULANT (G)": [
            "Fournisseurs débiteurs, avances et acomptes",
            "Clients et comptes rattachés",
            "État",
            "Comptes d’associés",
            "Autres débiteurs",
            "Comptes de régularisation - Actif"
        ],
        "TITRES ET VALEURS DE PLACEMENT (H)": [
            "Titres et valeurs de placement"
        ],
        "ÉCARTS DE CONVERSION – ACTIF (I)": [
            "Éléments circulants"
        ],
        "TRÉSORERIE – ACTIF (J)": [
            "Banques, T.G. et C.C.P.",
            "Chèques et valeurs à encaisser",
            "Caisse, régie d’avances et accréditifs"
        ]
    },
    "Bilan_Passif": {
        "CAPITAUX PROPRES": ["Capital social ou personnel", "Réserves", "Résultat net de l'exercice", "Autres capitaux propres"],
        "DETTES": ["Emprunts et dettes financières", "Fournisseurs et comptes rattachés", "Personnel", "État", "Autres dettes"],
        "PROVISIONS POUR RISQUES ET CHARGES": ["Provisions pour risques", "Provisions pour charges"],
        "TRÉSORERIE – PASSIF": ["Découverts bancaires", "Autres crédits de trésorerie"]
    }
}

def normalize_type_tableau(val):
    if val is None:
        return ""
    s = str(val).strip().lower()
    if "actif" in s:
        return "Bilan_Actif"
    if "passif" in s:
        return "Bilan_Passif"
    if "bilan" in s:
        return "Bilan_Actif"
    return str(val).strip()

def normalize_sous_categories(df):
    """Mappe les valeurs de Sous_Categorie extraites vers les clés EXACTES d'ALL_RUBRIQUES."""
    if df.empty:
        return df
    keys_actif = list(ALL_RUBRIQUES.get("Bilan_Actif", {}).keys())
    keys_passif = list(ALL_RUBRIQUES.get("Bilan_Passif", {}).keys())

    def norm(s):
        if s is None:
            return ""
        s0 = str(s).lower()
        s0 = re.sub(r'[\(\)\[\]\-–]', ' ', s0)
        s0 = re.sub(r'\s+', ' ', s0).strip()
        return s0

    def build_map(keys):
        m = {}
        for k in keys:
            m[norm(k)] = k
        return m

    map_act = build_map(keys_actif)
    map_pass = build_map(keys_passif)
    norm_act_keys = list(map_act.keys())
    norm_pass_keys = list(map_pass.keys())

    def match_best(s, type_tableau):
        s0 = norm(s)
        if s0 == "":
            return s
        if type_tableau and "passif" in str(type_tableau).lower():
            keys_norm = norm_pass_keys
            mapping = map_pass
        else:
            keys_norm = norm_act_keys
            mapping = map_act
        for kn in keys_norm:
            if s0 == kn or s0 in kn or kn in s0:
                return mapping[kn]
        cand = difflib.get_close_matches(s0, keys_norm, n=1, cutoff=0.55)
        if cand:
            return mapping[cand[0]]
        return s

    df["Sous_Categorie"] = df.apply(
        lambda row: match_best(row.get("Sous_Categorie", ""), row.get("Type_Tableau", "")),
        axis=1
    )
    return df

def fill_missing_rubriques(df, type_tableau="Bilan_Actif", fill_with_zero=False):
    filled_rows = []
    if type_tableau == "Bilan_Actif":
        ordre_sous_categories = [
            "IMMOBILISATIONS EN NON VALEUR (A)",
            "IMMOBILISATIONS INCORPORELLES (B)",
            "IMMOBILISATIONS CORPORELLES (C)",
            "IMMOBILISATIONS FINANCIÈRES (D)",
            "ÉCARTS DE CONVERSION – ACTIF (E)",
            "TOTAL I (A+B+C+D+E)",
            "STOCKS (F)",
            "CRÉANCES DE L'ACTIF CIRCULANT (G)",
            "TITRES ET VALEURS DE PLACEMENT (H)",
            "ÉCARTS DE CONVERSION – ACTIF (I)",
            "TOTAL II (F+G+H+I)",
            "TRÉSORERIE – ACTIF (J)",
            "TOTAL III (TRÉSORERIE – ACTIF)",
            "TOTAL GENERAL ACTIF"
        ]
    else:
        ordre_sous_categories = [
            "CAPITAUX PROPRES",
            "PROVISIONS POUR RISQUES ET CHARGES",
            "DETTES",
            "TRÉSORERIE – PASSIF",
            "TOTAL GENERAL PASSIF"
        ]

    default_val = 0.0 if fill_with_zero else pd.NA

    for sous_cat in ordre_sous_categories:
        rubriques_attendues = ALL_RUBRIQUES.get(type_tableau, {}).get(sous_cat, [])
        if not rubriques_attendues and "TOTAL" in sous_cat:
            rubriques_attendues = ["TOTAL"]

        df_sc = df[df["Sous_Categorie"] == sous_cat] if not df.empty else pd.DataFrame()

        if df_sc.empty:
            for rub in rubriques_attendues or [""]:
                filled_rows.append({
                    "Type_Tableau": type_tableau,
                    "Sous_Categorie": sous_cat,
                    "Rubrique": rub,
                    "Montant_Brut": default_val,
                    "Amortissements_Provisions": default_val,
                    "Net_Exercice": default_val,
                    "Net_Exercice_Prec": default_val,
                    "Commentaires": ""
                })
        else:
            rubriques_existantes = df_sc["Rubrique"].astype(str).tolist()
            for rub in rubriques_attendues or [""]:
                if rub in rubriques_existantes:
                    rows = df_sc[df_sc["Rubrique"] == rub].to_dict("records")
                    filled_rows.extend(rows)
                else:
                    filled_rows.append({
                        "Type_Tableau": type_tableau,
                        "Sous_Categorie": sous_cat,
                        "Rubrique": rub,
                        "Montant_Brut": default_val,
                        "Amortissements_Provisions": default_val,
                        "Net_Exercice": default_val,
                        "Net_Exercice_Prec": default_val,
                        "Commentaires": ""
                    })
    out_df = pd.DataFrame(filled_rows)
    cols = ["Type_Tableau", "Sous_Categorie", "Rubrique", "Montant_Brut", "Amortissements_Provisions", "Net_Exercice", "Net_Exercice_Prec", "Commentaires"]
    for c in cols:
        if c not in out_df.columns:
            out_df[c] = pd.NA
    return out_df[cols]

def clean_number_str(x):
    if x is None:
        return ""
    s = str(x).strip()
    if not s or s.lower() in ["nan", "none", "—", "-", "–", "--"]:
        return ""
    s = s.replace('\u00A0', '').replace(' ', '')
    s = re.sub(r'[^\d\-\.,\(\)]', '', s)
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    s = s.replace(',', '.')
    if s.count('.') > 1:
        parts = s.split('.')
        s = ''.join(parts[:-1]) + '.' + parts[-1]
    if re.match(r'^-?\d+(\.\d+)?$', s):
        try:
            return "{:.2f}".format(float(s))
        except:
            return ""
    return ""

# ----------------- Fonction principale de traitement -----------------
def ocr_pages_with_gemini(img_bytes_list, model_name="gemini-1.5-flash"):
    if 'PROMPT_BILAN_TABLEAU' not in globals():
        raise RuntimeError("PROMPT_BILAN_TABLEAU non défini.")
    batch_payload = [PROMPT_BILAN_TABLEAU]
    for b in img_bytes_list:
        batch_payload.append({"mime_type": "image/png", "data": base64.b64encode(b).decode("utf-8")})
    text = call_gemini(batch_payload, model_name=model_name)
    rows = safe_json_extract(text)
    if rows:
        # ensure keys exist consistently
        normalized = []
        for r in rows:
            normalized.append({k: r.get(k, None) for k in ["Type_Tableau","Sous_Categorie","Rubrique","Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec","Commentaires"]})
        return normalized
    if DEBUG:
        print("⚠️ safe_json_extract vide → tentative de parsing textuel de la réponse Gemini...")
    text_rows = parse_gemini_text_to_rows(text)
    if text_rows:
        txt_lower = text.lower()
        tt = None
        if 'actif' in txt_lower and 'passif' not in txt_lower:
            tt = "Bilan_Actif"
        elif 'passif' in txt_lower and 'actif' not in txt_lower:
            tt = "Bilan_Passif"
        for r in text_rows:
            r["Type_Tableau"] = r.get("Type_Tableau") or (tt if tt else "Bilan_Actif")
            for k in ["Montant_Brut", "Amortissements_Provisions", "Net_Exercice", "Net_Exercice_Prec"]:
                if r.get(k, "") == "":
                    r[k] = None
        if DEBUG:
            print(f"Parsed {len(text_rows)} rows from text fallback.")
        return text_rows

    fallback_rows = []
    for i, b in enumerate(img_bytes_list):
        if DEBUG:
            print(f"⚠️ Fallback page {i+1}/{len(img_bytes_list)} envoi séparé...")
        try:
            payload = [PROMPT_BILAN_TABLEAU, {"mime_type": "image/png", "data": base64.b64encode(b).decode("utf-8")}]
            t = call_gemini(payload, model_name=model_name)
            r = safe_json_extract(t)
            if r:
                fallback_rows.extend(r)
            else:
                pr = parse_gemini_text_to_rows(t)
                if pr:
                    fallback_rows.extend(pr)
        except Exception as e:
            print("Erreur sur page fallback:", e)
    return fallback_rows

def process_inputs(inputs, batch_size=5, out_xlsx=None, fill_with_zero=False):
    all_imgs = []
    if isinstance(inputs, str) and inputs.lower().endswith(".pdf"):
        if DEBUG:
            print("DEBUG: traitement comme PDF scanné")
        all_imgs = extract_images_from_pdf(inputs)
    elif isinstance(inputs, str) and os.path.isdir(inputs):
        for fname in sorted(os.listdir(inputs)):
            if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                with open(os.path.join(inputs, fname), "rb") as f:
                    all_imgs.append(f.read())
    elif isinstance(inputs, list):
        for img_path in inputs:
            with open(img_path, "rb") as f:
                all_imgs.append(f.read())
    else:
        raise ValueError("Entrée non reconnue. Donne un PDF, un dossier d'images ou une liste de fichiers image.")

    if not all_imgs:
        raise RuntimeError("⚠️ Aucune image trouvée pour traitement.")

    if out_xlsx is None:
        base = os.path.splitext(os.path.basename(inputs if isinstance(inputs, str) else "bilan"))[0]
        out_xlsx = f"result\\{base}_gemini.xlsx"

    all_rows = []
    for batch_start in range(0, len(all_imgs), batch_size):
        batch_end = min(batch_start + batch_size, len(all_imgs))
        print(f"📄 Traitement images {batch_start+1} à {batch_end} / {len(all_imgs)} ...")
        batch_imgs = all_imgs[batch_start:batch_end]
        rows = ocr_pages_with_gemini(batch_imgs)
        if rows:
            all_rows.extend(rows)

    if DEBUG:
        print("DEBUG - sample all_rows (premières lignes) :")
        for r in all_rows[:20]:
            print(r)

    df = pd.DataFrame(all_rows)
    if df.empty:
        print("⚠️ Aucun résultat exploitable extrait")
        return None

    # Standardiser noms colonnes
    df.columns = [str(c).strip() for c in df.columns]

    # Nettoyage rubriques
    for col in ["Sous_Categorie", "Rubrique"]:
        if col in df:
            df[col] = df[col].astype(str).str.strip()

    # Normaliser Type_Tableau
    if "Type_Tableau" in df:
        df["Type_Tableau"] = df["Type_Tableau"].apply(normalize_type_tableau)
    else:
        df["Type_Tableau"] = "Bilan_Actif"

    # Nettoyage valeurs numériques (string -> formatted -> float/NaN)
    for col in ["Montant_Brut", "Amortissements_Provisions", "Net_Exercice", "Net_Exercice_Prec"]:
        if col in df:
            df[col] = df[col].apply(lambda x: "" if x is None else x)
            df[col] = df[col].apply(clean_number_str)
            df[col] = pd.to_numeric(df[col].replace("", pd.NA), errors="coerce")

    # Supprimer lignes sans rubrique et sans montants
    df = df.dropna(how="all", subset=["Sous_Categorie", "Rubrique"] + ["Montant_Brut", "Amortissements_Provisions", "Net_Exercice", "Net_Exercice_Prec"])

    # --- NOUVEAU : normaliser les Sous_Categorie pour matcher ALL_RUBRIQUES keys ---
    if DEBUG:
        before = sorted(set(df.get("Sous_Categorie", pd.Series(dtype=str)).astype(str).tolist()))
        print("DEBUG - unique Sous_Categorie before normalizing (sample):", before[:50])
    df = normalize_sous_categories(df)
    if DEBUG:
        after = sorted(set(df.get("Sous_Categorie", pd.Series(dtype=str)).astype(str).tolist()))
        print("DEBUG - unique Sous_Categorie after normalizing (sample):", after[:50])

    # Séparer Actif / Passif
    df_actif = df[df["Type_Tableau"] == "Bilan_Actif"] if "Type_Tableau" in df else pd.DataFrame()
    df_passif = df[df["Type_Tableau"] == "Bilan_Passif"] if "Type_Tableau" in df else pd.DataFrame()

    # Remplir rubriques manquantes
    df_actif = fill_missing_rubriques(df_actif, "Bilan_Actif", fill_with_zero=fill_with_zero)
    df_passif = fill_missing_rubriques(df_passif, "Bilan_Passif", fill_with_zero=fill_with_zero)

    # Export Excel
    os.makedirs(os.path.dirname(out_xlsx), exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        if not df_actif.empty:
            df_actif.to_excel(writer, sheet_name="Actif", index=False)
        if not df_passif.empty:
            df_passif.to_excel(writer, sheet_name="Passif", index=False)

    print(f"✅ Extraction Bilan → {out_xlsx} ({len(df_actif)+len(df_passif)} lignes)")
    return out_xlsx

# ----------------- Entrée principale (exécution) -----------------
if __name__ == "__main__":
    pdf_path = "/mnt/data/TOLIMAR_2017.pdf"  # <-- change si besoin
    try:
        out = process_inputs(pdf_path, batch_size=3, fill_with_zero=False)
        print("Output:", out)
    except Exception as ex:
        print("Erreur pendant le traitement:", ex)
