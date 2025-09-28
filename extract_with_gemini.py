# -*- coding: utf-8 -*-
"""
extract_with_gemini_patched.py
Version: patch intégrant
 - mapping du JSON Gemini -> schéma interne attendu
 - fix normalize_sous_categories
 - debug prints supplémentaires
 - ajout prise en charge d'un 3ème tableau "Bilan_CPC" (template / patching / sheet Excel)

Ne nécessite pas pytesseract.
"""
import os
import re
import json
import base64
import difflib
import fitz  # PyMuPDF
import pandas as pd
from dotenv import load_dotenv
from pathlib import Path

# Gemini client (optionnel)
try:
    import google.generativeai as genai
except Exception:
    genai = None

# ----------------- Config / DEBUG -----------------
load_dotenv()
API_KEYS = [
    os.getenv("GEMINI_API_KEY1"),
    os.getenv("GEMINI_API_KEY2"),
    os.getenv("GEMINI_API_KEY3"),
]
API_KEYS = [k for k in API_KEYS if k]
current_key_index = 0
DEBUG = True

def init_genai():
    global current_key_index
    if not API_KEYS or genai is None:
        return
    genai.configure(api_key=API_KEYS[current_key_index])
    if DEBUG:
        print(f"🔑 GenAI key index: {current_key_index}")

def switch_key():
    global current_key_index
    if not API_KEYS or genai is None:
        return
    current_key_index = (current_key_index + 1) % len(API_KEYS)
    if DEBUG:
        print(f"⚠️ Switch to key index {current_key_index}")
    init_genai()

init_genai()

# --- Gemini: liste et sélection automatique du meilleur modèle ---
AVAILABLE_MODELS_MAP = {}
if genai:
    try:
        for m in genai.list_models():
            name = getattr(m, "name", None) or str(m)
            if name:
                clean = name.lower().strip()
                if clean.startswith("models/"):
                    clean = clean.split("/", 1)[1]
                AVAILABLE_MODELS_MAP[clean] = name
        if DEBUG:
            print("DEBUG modèles Gemini disponibles:", list(AVAILABLE_MODELS_MAP.keys()))
    except Exception as e:
        print("DEBUG: Impossible de lister les modèles Gemini:", e)
PREFERRED = ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"]
chosen_model_full = None
for pref in PREFERRED:
    if pref in AVAILABLE_MODELS_MAP:
        chosen_model_full = AVAILABLE_MODELS_MAP[pref]
        break
if not chosen_model_full and AVAILABLE_MODELS_MAP:
    chosen_model_full = next(iter(AVAILABLE_MODELS_MAP.values()))

# ---------------- Prompt (inchangé) -----------------
PROMPT_BILAN_TABLEAU = """
Tu es un parseur LLM spécialisé en bilans comptables français. Entrée : le texte OCR d'UNE SEULE PAGE (tableau scanné).
SORTIE : UNIQUEMENT un JSON (une liste d'objets) — rien d'autre, pas de texte explicatif, pas de balises Markdown.

FORMAT EXACT attendu (chaque item doit contenir ces clés EXACTES) :
[ ... voir ton prompt original ... ]
"""
PROMPT_PAGE_ONLY = PROMPT_BILAN_TABLEAU

_re_number = re.compile(r'[-]?\(?\d{1,3}(?:[ \u00A0\d]{0,}\d)?[.,]\d{1,2}\)?|-?\d+')

# ---------------- utilitaires nombres ----------------
def best_match_name(lib, candidats, cutoff=0.55):
    """
    Retourne le meilleur candidat (str) dont la similarité avec lib dépasse cutoff.
    Utilise difflib.get_close_matches.
    """
    from difflib import get_close_matches
    lib_norm = normalize_str_for_match(lib)
    candidats_norm = [normalize_str_for_match(c) for c in candidats]
    matches = get_close_matches(lib_norm, candidats_norm, n=1, cutoff=cutoff)
    if matches:
        idx = candidats_norm.index(matches[0])
        return candidats[idx]
    return None

def normalize_str_for_match(s):
    if s is None:
        return ""
    s = str(s).lower().strip()
    s = re.sub(r'[\s\-_]+', ' ', s)
    s = re.sub(r'[^\w\s]', '', s)
    # simplifications accents
    s = s.replace('é', 'e').replace('è', 'e').replace('ê', 'e').replace('à', 'a').replace('ç', 'c')
    return s

def normalize_number_str(s):
    if s is None:
        return ""
    s = str(s).strip()
    if s in ["", "-", "—", "–", "--", "nan", "None"]:
        return ""
    s = s.replace('\u00A0', ' ')
    m = _re_number.search(s)
    if not m:
        s2 = re.sub(r'[^\d\-,\.]', '', s)
        if not s2:
            return ""
        s = s2
    else:
        s = m.group(0)
    s = s.replace(' ', '')
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

# ---------------- Hierarchical template helpers ----------------
def pretty_name(key):
    if not key:
        return ""
    s = str(key)
    s = s.replace('_', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s.upper()

def create_flat_template_from_hier(tpl_hier, type_tableau="Bilan_Actif"):
    out = []
    side = tpl_hier.get(type_tableau, {})
    sous_cats = side.get("Sous_Categories", {}) if isinstance(side, dict) else {}
    for group_name, group_data in sous_cats.items():
        rubs = group_data.get("Rubriques", {})
        parent_key_for_group = str(group_name).strip()
        for parent_key, children in rubs.items():
            sous_label = pretty_name(parent_key)
            if isinstance(children, list):
                for child in children:
                    out.append({
                        "Type_Tableau": type_tableau,
                        "Parent_Sous_Categorie": parent_key_for_group,
                        "Sous_Categorie": sous_label,
                        "Rubrique": str(child).strip(),
                        "Montant_Brut": "",
                        "Amortissements_Provisions": "",
                        "Net_Exercice": "",
                        "Net_Exercice_Prec": "",
                        "Commentaires": ""
                    })
            else:
                out.append({
                    "Type_Tableau": type_tableau,
                    "Parent_Sous_Categorie": parent_key_for_group,
                    "Sous_Categorie": sous_label,
                    "Rubrique": sous_label if isinstance(children, str) else str(parent_key),
                    "Montant_Brut": "",
                    "Amortissements_Provisions": "",
                    "Net_Exercice": "",
                    "Net_Exercice_Prec": "",
                    "Commentaires": "_is_total_template"
                })
    return out

def load_template_json(path):
    """
    Lit un template JSON. Supporte :
     - liste (flat) -> retourne liste, numeric_keys détectés dynamiquement
     - dictionnaire hiérarchique contenant Bilan_Actif ou Bilan_Passif -> aplatie
    Remarque : si c'est une liste arbitraire (ex: Bilan_CPC.json fourni par l'utilisateur), on retourne tel quel.
    """
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    if isinstance(data, list):
        sample = data[0] if data else {}
        numeric_keys = [k for k in sample.keys() if k not in ("Type_Tableau","Parent_Sous_Categorie","Sous_Categorie","Rubrique","Commentaires")]
        return data, numeric_keys
    if isinstance(data, dict):
        # si c'est une structure hiérarchique, chercher la première clé Bilan_* présente
        for possible in ("Bilan_Actif", "Bilan_Passif", "Bilan_CPC"):
            if possible in data:
                flat = create_flat_template_from_hier(data, possible)
                sample = flat[0] if flat else {}
                numeric_keys = [k for k in sample.keys() if k not in ("Type_Tableau","Parent_Sous_Categorie","Sous_Categorie","Rubrique","Commentaires")]
                return flat, numeric_keys
        # fallback: erreur
    raise ValueError("Template JSON non reconnu (attendu liste ou structure hiérarchique).")

def build_rubrique_to_sous_map(flat_template):
    rub_to_sous = {}
    rub_to_parent = {}
    candidats = []
    for row in flat_template:
        rub = str(row.get("Rubrique","")).strip()
        sous = str(row.get("Sous_Categorie","")).strip()
        parent = str(row.get("Parent_Sous_Categorie","")).strip() if row.get("Parent_Sous_Categorie") else ""
        if rub:
            key = normalize_str_for_match(rub)
            rub_to_sous[key] = sous
            rub_to_parent[key] = parent
            candidats.append(rub)
    candidats = list(dict.fromkeys(candidats))
    return rub_to_sous, rub_to_parent, candidats

# ----------------- parsing helpers (inchangés / améliorés) -----------------
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

# ---------------- Fallback number-finder ----------------
def find_numbers_near_label(page_text, label, window_lines=2):
    if not page_text or not label:
        return []
    lines = [ln for ln in page_text.splitlines()]
    lab_norm = normalize_str_for_match(label)
    nums = []
    for idx, ln in enumerate(lines):
        if lab_norm and lab_norm in normalize_str_for_match(ln):
            found = _re_number.findall(ln)
            if found:
                nums.extend([normalize_number_str(x) for x in found if normalize_number_str(x)])
            for j in range(1, window_lines+1):
                if idx + j < len(lines):
                    found2 = _re_number.findall(lines[idx+j])
                    if found2:
                        nums.extend([normalize_number_str(x) for x in found2 if normalize_number_str(x)])
                if idx - j >= 0:
                    found3 = _re_number.findall(lines[idx-j])
                    if found3:
                        nums.extend([normalize_number_str(x) for x in found3 if normalize_number_str(x)])
            break
    out = []
    for n in nums:
        if n and n not in out:
            out.append(n)
    return out

# -------------- template creation / ensure (modifié) ------------------
def ensure_templates_exist(tpl_act_path="templates/bilan_actif.json",
                           tpl_pass_path="templates/bilan_passif.json",
                           tpl_cpc_path="templates/bilan_cpc.json"):
    os.makedirs(os.path.dirname(tpl_act_path) or ".", exist_ok=True)
    os.makedirs(os.path.dirname(tpl_pass_path) or ".", exist_ok=True)
    os.makedirs(os.path.dirname(tpl_cpc_path) or ".", exist_ok=True)

    if not os.path.exists(tpl_act_path):
        sample = [
            {
                "Type_Tableau": "Bilan_Actif",
                "Parent_Sous_Categorie": "Actif_Immobilisation",
                "Sous_Categorie": "IMMOBILISATIONS EN NON VALEUR (A)",
                "Rubrique": "Frais préliminaires",
                "Montant_Brut": "",
                "Amortissements_Provisions": "",
                "Net_Exercice": "",
                "Net_Exercice_Prec": "",
                "Commentaires": ""
            }
        ]
        with open(tpl_act_path, "w", encoding="utf-8") as f:
            json.dump(sample, f, ensure_ascii=False, indent=2)
        if DEBUG: print(f"⚠️ Template actif manquant -> fichier placeholder créé : {tpl_act_path}")
    if not os.path.exists(tpl_pass_path):
        sample = [
            {
                "Type_Tableau": "Bilan_Passif",
                "Parent_Sous_Categorie": "Passif_Capitaux",
                "Sous_Categorie": "CAPITAUX PROPRES (A)",
                "Rubrique": "Capital social ou personnel (1)",
                "Montant_Brut": "",
                "Amortissements_Provisions": "",
                "Net_Exercice": "",
                "Net_Exercice_Prec": "",
                "Commentaires": ""
            }
        ]
        with open(tpl_pass_path, "w", encoding="utf-8") as f:
            json.dump(sample, f, ensure_ascii=False, indent=2)
        if DEBUG: print(f"⚠️ Template passif manquant -> fichier placeholder créé : {tpl_pass_path}")
    if not os.path.exists(tpl_cpc_path):
        # Sample minimal CPC template (list style)
        sample_cpc = [
            {
                "Type_Tableau": "Bilan_CPC",
                "Parent_Sous_Categorie": "Exploitation",
                "Sous_Categorie": "PRODUITS D'EXPLOITATION",
                "Rubrique": "PRODUITS D'EXPLOITATION",
                "Propres a l'exercice": "",
                "Concernant les exercices_prec": "",
                "Totaux de l'exercice": "",
                "Exercice_prec": "",
                "Commentaires": ""
            }
        ]
        with open(tpl_cpc_path, "w", encoding="utf-8") as f:
            json.dump(sample_cpc, f, ensure_ascii=False, indent=2)
        if DEBUG: print(f"⚠️ Template CPC manquant -> fichier placeholder créé : {tpl_cpc_path}")

# -------------- call_gemini (tolérant) -------------------
def call_gemini(content, model_name="gemini-2.5-flash"):
    init_genai()
    if not API_KEYS or genai is None:
        if DEBUG: print("⚠️ Pas de clé Gemini ou lib manquante → on saute l'appel LLM (retour vide).")
        return ""
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
        if DEBUG: print("Erreur Gemini capturée :", repr(e))
        err = str(e).lower()
        if "429" in err or "quota" in err or "rate limit" in err:
            try:
                switch_key()
                return call_gemini(content, model_name)
            except Exception as ee:
                if DEBUG: print("Echec bascule clé :", ee)
                return ""
        return ""

# -------------- extraction helpers -----------------------
def extract_text_from_pdf_page(page):
    try:
        return page.get_text("text")
    except Exception:
        return ""

def page_to_base64_image_bytes(page, zoom=2):
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return pix.tobytes("png")

TECH_COLS = ["_matched_from_page", "_matched_name_raw", "_source"]

def ensure_tech_cols(df):
    for col in TECH_COLS:
        if col not in df.columns:
            df[col] = None
    return df

# -------------- LLM JSON -> internal mapping (NEW) --------------
def map_gemini_json_to_internal(rows):
    if not rows or not isinstance(rows, list):
        return rows
    out = []
    key_candidates = {
        "rubrique": ["rubrique", "elements", "libellé", "éléments", "label", "elements"],
        "brut": ["brut", "brut ", "brut€", "brut_value"],
        "amort": ["amort. & prov.", "amort.&prov.", "amort. & prov", "amortissements & provisions", "amort", "amortissements"],
        "net": ["net", "net ", "net_exercice"],
        "prev": ["exercice précédent", "exercice_prec", "exercice précédent ", "exercice_precedent", "exercice précédent"]
    }

    def find_key(d, candidates):
        for k in d.keys():
            kl = str(k).lower().strip()
            for cand in candidates:
                if cand in kl:
                    return k
        return None

    for item in rows:
        if not isinstance(item, dict):
            continue
        new = {
            "Type_Tableau": "",
            "Sous_Categorie": "",
            "Rubrique": "",
            "Montant_Brut": "",
            "Amortissements_Provisions": "",
            "Net_Exercice": "",
            "Net_Exercice_Prec": "",
            "Commentaires": ""
        }
        k_r = find_key(item, key_candidates["rubrique"])
        k_b = find_key(item, key_candidates["brut"])
        k_a = find_key(item, key_candidates["amort"])
        k_n = find_key(item, key_candidates["net"])
        k_p = find_key(item, key_candidates["prev"])

        if not k_r:
            for k,v in item.items():
                if isinstance(v, str) and re.search(r'[A-Za-zÀ-ÖØ-öø-ÿ]', v):
                    k_r = k
                    break

        if k_r:
            new["Rubrique"] = str(item.get(k_r) or "").strip()
        if k_b:
            new["Montant_Brut"] = normalize_number_str(item.get(k_b))
        if k_a:
            new["Amortissements_Provisions"] = normalize_number_str(item.get(k_a))
        if k_n:
            new["Net_Exercice"] = normalize_number_str(item.get(k_n))
        if k_p:
            new["Net_Exercice_Prec"] = normalize_number_str(item.get(k_p))

        new["_raw"] = item
        out.append(new)
    return out

# -------------- normalize_sous_categories & fill_missing_rubriques (fixed) --------------
def normalize_sous_categories(df):
    def infer_parent(sous):
        if not sous: return ""
        s = str(sous).upper()
        if "IMMOBILISATION" in s:
            return "Actif_Immobilisation"
        if "STOCK" in s or "CRÉANCE" in s or "CREANCE" in s or "COMPTES" in s:
            return "Actif_Circulant"
        if "TRÉSORERIE" in s or "TRESORERIE" in s or "BANQUE" in s:
            return "Tresorerie"
        return ""
    if "Parent_Sous_Categorie" not in df.columns:
        df["Parent_Sous_Categorie"] = ""
    df["Parent_Sous_Categorie"] = df["Parent_Sous_Categorie"].fillna("")
    for idx, row in df.iterrows():
        if not row.get("Parent_Sous_Categorie"):
            df.at[idx, "Parent_Sous_Categorie"] = infer_parent(row.get("Sous_Categorie",""))
    return df

def fill_missing_rubriques(df, type_tableau, fill_with_zero=False):
    # For CPC templates numeric fields may have different names; this helper keeps original behaviour for Bilan_Actif
    if fill_with_zero and isinstance(df, pd.DataFrame):
        numeric_cols = ["Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec"]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: "0.00" if (x in (None,"") or (isinstance(x, float) and pd.isna(x))) else x)
    return df

# -------------- Main processing (template-first) ----------------
def process_pdf_with_templates(pdf_path,
                               tpl_act_path="templates/bilan_actif.json",
                               tpl_pass_path="templates/bilan_passif.json",
                               tpl_cpc_path="templates/bilan_cpc.json",
                               use_gemini=True, reverse_numbers=False, out_xlsx=None, fill_with_zero=False, zoom=2):
    ensure_templates_exist(tpl_act_path, tpl_pass_path, tpl_cpc_path)

    tpl_act, num_keys_act = load_template_json(tpl_act_path)
    tpl_pass, num_keys_pass = load_template_json(tpl_pass_path)
    tpl_cpc, num_keys_cpc = load_template_json(tpl_cpc_path)

    rub_to_sous_act, rub_to_parent_act, candidats_act = build_rubrique_to_sous_map(tpl_act)
    rub_to_sous_pass, rub_to_parent_pass, candidats_pass = build_rubrique_to_sous_map(tpl_pass)
    rub_to_sous_cpc,  rub_to_parent_cpc,  candidats_cpc  = build_rubrique_to_sous_map(tpl_cpc)

    if DEBUG:
        print("Templates chargés -> Actif:", len(tpl_act), "Passif:", len(tpl_pass), "CPC:", len(tpl_cpc))
        print("Candidats Actif exemples:", candidats_act[:8])
        print("Candidats CPC exemples:", candidats_cpc[:8])

    doc = fitz.open(pdf_path)
    patched_act = []
    patched_pass = []
    patched_cpc = []

    for p_index, page in enumerate(doc):
        if DEBUG: print(f"--- page {p_index+1}/{len(doc)} ---")
        page_text = extract_text_from_pdf_page(page) or ""
        payload = None
        if page_text and len(page_text.strip()) > 100:
            payload = [PROMPT_PAGE_ONLY, page_text]
        else:
            if not use_gemini:
                if DEBUG: print("Page scannée et use_gemini=False -> skip page.")
                continue
            img_bytes = page_to_base64_image_bytes(page, zoom=zoom)
            b64 = base64.b64encode(img_bytes).decode('utf-8')
            payload = [PROMPT_PAGE_ONLY, {"mime_type":"image/png", "data": b64}]

        raw_text = ""
        if use_gemini:
            try:
                raw_text = call_gemini(payload, model_name=chosen_model_full)
            except Exception as e:
                if DEBUG: print("Gemini call exception:", e)
                raw_text = ""
        else:
            raw_text = ""

        rows = []
        if raw_text:
            rows = safe_json_extract(raw_text)
            # Si Gemini a renvoyé une liste d'objets, mappez-les vers le schéma interne
            if rows and isinstance(rows, list) and isinstance(rows[0], dict):
                rows = map_gemini_json_to_internal(rows)
            if not rows:
                rows = parse_gemini_text_to_rows(raw_text)
        else:
            if page_text and len(page_text.strip()) > 50:
                if DEBUG: print("Gemini absent -> parsing local du texte de la page.")
                rows = parse_gemini_text_to_rows(page_text)
            else:
                rows = []

        if DEBUG:
            print("DEBUG parsed rows count:", len(rows))
            if len(rows) <= 20:
                print("DEBUG rows preview:")
                for r in rows[:20]:
                    print(r)

        def postprocess_llm_rows(rows):
            subtotal_map = {}
            for r in rows:
                if r.get("is_subcategory") and r.get("subtotal_cells"):
                    brut = r["subtotal_cells"][0] if r["subtotal_cells"] else ""
                    if brut:
                        subtotal_map[normalize_str_for_match(r.get("Sous_Categorie",""))] = normalize_number_str(brut)
            cleaned = []
            subtotals = []
            for r in rows:
                if r.get("Commentaires") == "_is_subtotal":
                    subtotals.append(r)
                    continue
                if not r.get("is_subcategory"):
                    key = normalize_str_for_match(r.get("Sous_Categorie",""))
                    brut = normalize_number_str(r.get("Montant_Brut",""))
                    if key in subtotal_map and brut == subtotal_map[key]:
                        r["Montant_Brut"] = ""
                        r["Amortissements_Provisions"] = ""
                        r["Commentaires"] = (r.get("Commentaires","") + " _cleared_equal_subtotal").strip()
                for s in subtotals:
                    if r.get("Sous_Categorie") == s.get("Sous_Categorie"):
                        if r.get("Amortissements_Provisions") == s.get("Amortissements_Provisions") and r.get("Amortissements_Provisions"):
                            r["Amortissements_Provisions"] = ""
                cleaned.append(r)
            return cleaned, subtotals

        rows_raw = rows
        rows, subtotals = postprocess_llm_rows(rows_raw)

        # ---------- Matching robuste : départage Actif vs Passif vs CPC ----------
        for r in rows:
            lib = str(r.get("Rubrique","")).strip()
            if not lib:
                continue

            lib_norm = normalize_str_for_match(lib)

            # Meilleur candidat côté passif, actif, cpc
            matched_pass = best_match_name(lib, candidats_pass, cutoff=0.55)
            matched_act  = best_match_name(lib, candidats_act,  cutoff=0.55)
            matched_cpc  = best_match_name(lib, candidats_cpc,  cutoff=0.55)

            # Similarité pour départager
            from difflib import SequenceMatcher
            def similarity(a, b):
                if not a or not b:
                    return 0.0
                return SequenceMatcher(None, normalize_str_for_match(a), normalize_str_for_match(b)).ratio()

            score_pass = similarity(lib, matched_pass) if matched_pass else 0.0
            score_act  = similarity(lib, matched_act)  if matched_act  else 0.0
            score_cpc  = similarity(lib, matched_cpc)  if matched_cpc  else 0.0

            if DEBUG:
                print(f"[MATCH DEBUG] '{lib[:80]}' -> passif:{matched_pass} (s{score_pass:.3f}) | actif:{matched_act} (s{score_act:.3f}) | cpc:{matched_cpc} (s{score_cpc:.3f})")

            # Règles de décision multi-côtés
            threshold = 0.45   # score min pour accepter un match
            # on prend le max score si suffisamment au-dessus des autres
            scores = {"passif": score_pass, "actif": score_act, "cpc": score_cpc}
            best_side = max(scores, key=scores.get)
            best_score = scores[best_side]
            # second best
            sorted_scores = sorted(scores.items(), key=lambda x: x[1], reverse=True)
            second_best_score = sorted_scores[1][1] if len(sorted_scores) > 1 else 0.0
            delta = 0.07

            chosen_side = None
            chosen_matched = None
            if best_score >= threshold and (best_score - second_best_score > delta):
                chosen_side = best_side
            else:
                # scores trop proches -> regarder le contexte page (mot 'actif'/'passif'/'produit'/'charges' etc)
                ptxt = (page_text or "").lower()
                if "passif" in ptxt and "actif" not in ptxt:
                    chosen_side = "passif"
                elif "actif" in ptxt and "passif" not in ptxt:
                    chosen_side = "actif"
                else:
                    # essayer détecter mot-cle CPC (produit/charge/chiffre d'affaire)
                    if any(k in lib.lower() for k in ["vente", "produit", "charge", "exploitation", "chiffre"]):
                        chosen_side = "cpc"
                    else:
                        # fallback: côté meilleur score même si proche
                        chosen_side = best_side

            if chosen_side == "passif":
                chosen_matched = matched_pass
            elif chosen_side == "actif":
                chosen_matched = matched_act
            else:
                chosen_matched = matched_cpc

            if not chosen_matched:
                if DEBUG:
                    print("Aucun candidat choisi pour:", lib)
                continue

            # Sélection du template / keys selon le côté choisi
            if chosen_side == "passif":
                tpl_list = tpl_pass
                nkeys = num_keys_pass
                rub_to_parent = rub_to_parent_pass
            elif chosen_side == "actif":
                tpl_list = tpl_act
                nkeys = num_keys_act
                rub_to_parent = rub_to_parent_act
            else:
                tpl_list = tpl_cpc
                nkeys = num_keys_cpc
                rub_to_parent = rub_to_parent_cpc

            keym = normalize_str_for_match(chosen_matched)
            tpl_entries = [t for t in tpl_list if normalize_str_for_match(t.get("Rubrique","")) == keym]
            if not tpl_entries:
                tpl_entries = [t for t in tpl_list if t.get("Rubrique","") == chosen_matched]
            if not tpl_entries:
                if DEBUG:
                    print("Template entry introuvable malgré match -> skip", chosen_matched)
                continue

            # Récupération / normalisation des nombres (mapping spécifique à chaque bilan)
            if chosen_side == "actif":
                mapped_fields = [
                    r.get("Montant_Brut", ""),
                    r.get("Amortissements_Provisions", ""),
                    r.get("Net_Exercice", ""),
                    r.get("Net_Exercice_Prec", "")
                ]
            elif chosen_side == "passif":
                mapped_fields = [
                    r.get("Net_Exercice", ""),
                    r.get("Net_Exercice_Prec", "")
                ]
            elif chosen_side == "cpc":
                mapped_fields = [
                    r.get("Propres a l'exercice", ""),
                    r.get("Concernant les exercices_prec", ""),
                    r.get("Totaux de l'exercice", ""),
                    r.get("Exercice_prec", "")
                ]
            else:
                mapped_fields = []

            mapped_nums = [normalize_number_str(x) for x in mapped_fields if x not in (None,"") and normalize_number_str(x) != ""]

            # Initialisation par défaut
            nums_norm = []

            # ... mapping spécifique à chaque bilan ...
            if mapped_nums:
                nums_norm = mapped_nums
                if DEBUG:
                    print("Using mapped numeric fields for:", lib, "->", nums_norm)
            else:
                nums = []
                for v in r.values():
                    sv = ""
                    try:
                        sv = str(v)
                    except:
                        sv = ""
                    if re.search(r'\d', sv):
                        found = _re_number.findall(sv)
                        for f in found:
                            if f not in nums:
                                nums.append(f)
                if not nums:
                    line_text = " ".join([str(v) for v in r.values() if isinstance(v, str)])
                    found = _re_number.findall(line_text)
                    nums = found
                nums_norm = [normalize_number_str(n) for n in nums if normalize_number_str(n) != ""]
                if not nums_norm:
                    nums_near = find_numbers_near_label(page_text, lib, window_lines=2)
                    if nums_near:
                        if DEBUG:
                            print("Fallback numbers found near label:", lib, "->", nums_near)
                        nums_norm = nums_near

            if reverse_numbers:
                nums_norm = list(reversed(nums_norm))

            aligned_vals = align_numbers(nums_norm, nkeys)
            newrow = tpl_entries[0].copy()
            for i, nk in enumerate(nkeys):
                val = aligned_vals[i]
                if val not in (None, ""):
                    newrow[nk] = val
            if r.get("Sous_Categorie"):
                newrow["Sous_Categorie"] = r.get("Sous_Categorie")

            # Parent selon côté
            if not newrow.get("Parent_Sous_Categorie"):
                parent_mapped = rub_to_parent.get(keym)
                if parent_mapped:
                    newrow["Parent_Sous_Categorie"] = parent_mapped

            newrow["_matched_from_page"] = p_index+1
            newrow["_matched_name_raw"] = lib
            newrow["_source"] = "gemini_page"

            if chosen_side == "passif":
                patched_pass.append(newrow)
            elif chosen_side == "actif":
                patched_act.append(newrow)
            else:
                patched_cpc.append(newrow)

    # DEBUG print pour vérification
    if DEBUG:
        print(">>> patched_act count:", len(patched_act))
        print(">>> patched_pass count:", len(patched_pass))
        print(">>> patched_cpc count:", len(patched_cpc))
        if patched_act:  print("patched_act sample:", patched_act[:5])
        if patched_pass: print("patched_pass sample:", patched_pass[:5])
        if patched_cpc:  print("patched_cpc sample:", patched_cpc[:5])

    doc.close()

    if DEBUG:
        print("patched_act sample:", patched_act[:8])
        print("patched_pass sample:", patched_pass[:8])
        print("patched_cpc sample:", patched_cpc[:8])

    def merge_template_with_patches(template_list, patched_rows, numeric_keys):
        out = []
        patch_map = {}
        for p in patched_rows:
            key = normalize_str_for_match(p.get("Rubrique",""))
            # keep first patch encountered (last one wins if repeated)
            patch_map[key] = p
        for t in template_list:
            keyt = normalize_str_for_match(t.get("Rubrique",""))
            if keyt in patch_map:
                p = patch_map[keyt]
                row = t.copy()
                for nk in numeric_keys:
                    val = p.get(nk)
                    if val not in (None, "") and not (isinstance(val, type(pd.NA)) and pd.isna(val)):
                        row[nk] = val
                row["_matched_from_page"] = p.get("_matched_from_page")
                row["_matched_name_raw"] = p.get("_matched_name_raw")
                row["_source"] = p.get("_source", "")
                out.append(row)
            else:
                out.append(t.copy())
        return out

    # === Merge templates with patches ===
    merged_pass = merge_template_with_patches(tpl_pass, patched_pass, num_keys_pass)
    merged_act  = merge_template_with_patches(tpl_act, patched_act, num_keys_act)
    merged_cpc  = merge_template_with_patches(tpl_cpc, patched_cpc, num_keys_cpc)

    df_act = pd.DataFrame(merged_act) if merged_act else pd.DataFrame(tpl_act)
    df_pass = pd.DataFrame(merged_pass) if merged_pass else pd.DataFrame(tpl_pass)
    df_cpc = pd.DataFrame(merged_cpc) if merged_cpc else pd.DataFrame(tpl_cpc)

    df_act = df_act.fillna("")
    df_pass = df_pass.fillna("")
    df_cpc = df_cpc.fillna("")

    # Normalisation colonnes numériques (sauf colonnes matched/raw/source)
    for df in (df_act, df_pass, df_cpc):
        if df is None or df.empty:
            continue
        for col in df.columns:
            if col in ("Type_Tableau","Parent_Sous_Categorie","Sous_Categorie","Rubrique","Commentaires","_matched_from_page","_matched_name_raw","_source"):
                continue
            try:
                df[col] = df[col].apply(lambda x: normalize_number_str(x) if x not in (None,"") else x)
            except Exception:
                pass

    try:
        if not df_act.empty:
            df_act["Type_Tableau"] = df_act.get("Type_Tableau", "Bilan_Actif")
            if "Sous_Categorie" in df_act.columns:
                df_act = normalize_sous_categories(df_act)
            df_act = ensure_tech_cols(df_act)

        if not df_pass.empty:
            df_pass["Type_Tableau"] = df_pass.get("Type_Tableau", "Bilan_Passif")
            if "Sous_Categorie" in df_pass.columns:
                df_pass = normalize_sous_categories(df_pass)
            df_pass = ensure_tech_cols(df_pass)

        if not df_cpc.empty:
            df_cpc["Type_Tableau"] = df_cpc.get("Type_Tableau", "Bilan_CPC")
            df_cpc = ensure_tech_cols(df_cpc)
    except Exception as e:
        if DEBUG: print("Warning normalisation Sous_Categorie:", e)

    try:
        df_act = df_act.fillna("")
        df_pass = df_pass.fillna("")
        df_cpc = df_cpc.fillna("")

        if not df_act.empty and set(["Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec"]).issubset(set(df_act.columns)):
            cols_keep = [
                "Type_Tableau","Parent_Sous_Categorie","Sous_Categorie","Rubrique",
                "Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec",
                "Commentaires","_matched_from_page","_matched_name_raw","_source"
            ]
            cols_present = [c for c in cols_keep if c in df_act.columns]
            df_act2 = df_act[cols_present].copy()
            df_act = fill_missing_rubriques(df_act2, "Bilan_Actif", fill_with_zero=fill_with_zero)
    except Exception as e:
        if DEBUG: print("Warning fill_missing_rubriques:", e)

    if out_xlsx is None:
        base = Path(pdf_path).stem
        out_xlsx = f"result/{base}_patched_templates.xlsx"

    out_xlsx = os.path.abspath(out_xlsx)
    out_dir = os.path.dirname(out_xlsx)
    os.makedirs(out_dir, exist_ok=True)

    excel_written = False
    try:
        # -> Construire un dictionnaire sheet_name -> dataframe
        sheets = {}

        # On privilégie df_act / df_pass / df_cpc s'ils existent
        if df_act is not None and not df_act.empty:
            sheets["Bilan_Actif"] = df_act
        elif tpl_act:
            sheets["Bilan_Actif"] = pd.DataFrame(tpl_act)

        if df_pass is not None and not df_pass.empty:
            sheets["Bilan_Passif"] = df_pass
        elif tpl_pass:
            sheets["Bilan_Passif"] = pd.DataFrame(tpl_pass)

        if df_cpc is not None and not df_cpc.empty:
            sheets["Bilan_CPC"] = df_cpc
        elif tpl_cpc:
            sheets["Bilan_CPC"] = pd.DataFrame(tpl_cpc)

        # Au cas où il y aurait d'autres Type_Tableau dans merged_* non couverts :
        def add_by_type(df, default_name_prefix="Feuille"):
            if df is None or df.empty:
                return
            if "Type_Tableau" in df.columns:
                for ttype, sub in df.groupby("Type_Tableau"):
                    name = str(ttype).strip() or default_name_prefix
                    if name not in sheets:
                        sheets[name] = sub.reset_index(drop=True)

        add_by_type(df_act)
        add_by_type(df_pass)
        add_by_type(df_cpc)

        # Si pour une quelconque raison sheets est vide (fallback)
        if not sheets:
            sheets["Actif"] = pd.DataFrame(tpl_act) if tpl_act else pd.DataFrame()
            sheets["Passif"] = pd.DataFrame(tpl_pass) if tpl_pass else pd.DataFrame()
            sheets["CPC"] = pd.DataFrame(tpl_cpc) if tpl_cpc else pd.DataFrame()

        # Ecrire chaque dataframe dans sa propre feuille
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            for sheet_name, df_sheet in sheets.items():
                safe_name = sheet_name[:31]
                if df_sheet is None:
                    df_sheet = pd.DataFrame()
                try:
                    df_sheet.to_excel(writer, sheet_name=safe_name, index=False)
                except Exception as e:
                    alt = safe_name[:25] + "_1"
                    if DEBUG: print(f"Warning écriture feuille {safe_name} failed -> try {alt} :", e)
                    df_sheet.to_excel(writer, sheet_name=alt, index=False)

        excel_written = os.path.isfile(out_xlsx)
        if excel_written and DEBUG:
            print("✅ Excel écrit (feuilles):", out_xlsx, "sheets:", list(sheets.keys()))
    except Exception as e:
        if DEBUG:
            print("❌ Erreur écriture Excel :", repr(e))
        try:
            csv_act = os.path.join(out_dir, Path(out_xlsx).stem + "_Actif.csv")
            csv_pass = os.path.join(out_dir, Path(out_xlsx).stem + "_Passif.csv")
            csv_cpc = os.path.join(out_dir, Path(out_xlsx).stem + "_CPC.csv")
            if DEBUG:
                print("Tentative écriture CSV de secours :", csv_act, csv_pass, csv_cpc)
            if df_act is not None and not df_act.empty:
                df_act.to_csv(csv_act, index=False, sep=";")
            if df_pass is not None and not df_pass.empty:
                df_pass.to_csv(csv_pass, index=False, sep=";")
            if df_cpc is not None and not df_cpc.empty:
                df_cpc.to_csv(csv_cpc, index=False, sep=";")
            excel_written = os.path.isfile(csv_act) or os.path.isfile(csv_pass) or os.path.isfile(csv_cpc)
        except Exception as e:
            if DEBUG:
                print("❌ Erreur écriture CSV de secours :", repr(e))

    if DEBUG:
        print("Traitement terminé. Fichier de sortie :", out_xlsx)

    return out_xlsx

def safe_json_extract(text):
    """
    Essaie d'extraire une liste d'objets JSON depuis un texte Gemini.
    Retourne [] si rien n'est extrait.
    """
    import json
    try:
        text = text.strip()
        if text.startswith("```json"):
            text = text[7:]
        if text.endswith("```"):
            text = text[:-3]
        # Parfois Gemini renvoie du texte + JSON, tenter d'extraire le premier tableau JSON
        match = re.search(r'(\[.*\])', text, re.S)
        if match:
            raw = match.group(1)
            data = json.loads(raw)
            return data if isinstance(data, list) else [data]
        # fallback: essayer de charger tout
        obj = json.loads(text)
        if isinstance(obj, list):
            return obj
        if isinstance(obj, dict):
            return [obj]
        return []
    except Exception:
        # fallback permissif: retour vide si pas JSON valide
        if DEBUG:
            print("safe_json_extract: impossible d'extraire JSON. snippet:", text[:200])
        return []

# ------------- wrapper utilisé par Flask ----------------
def process_inputs(pdf_path, out_xlsx=None,
                  tpl_act="templates/bilan_actif.json",
                  tpl_pass="templates/bilan_passif.json",
                  tpl_cpc="templates/bilan_cpc.json",
                  client_id=None): 
    ensure_templates_exist(tpl_act, tpl_pass, tpl_cpc)

    result = process_pdf_with_templates(
        pdf_path,
        tpl_act_path=tpl_act,
        tpl_pass_path=tpl_pass,
        tpl_cpc_path=tpl_cpc,
        use_gemini=True,
        reverse_numbers=False,
        out_xlsx=out_xlsx,
        fill_with_zero=False,
        zoom=2
    )

    print(f"DEBUG process_inputs: result = {result}")
    print(f"DEBUG process_inputs: result exists? {os.path.exists(result) if result else 'result is None'}")

    return result

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Traitement PDF -> Excel (template-first).")
    parser.add_argument("--pdf", "-p", help="Chemin vers le PDF à traiter", required=False)
    parser.add_argument("--tpl_act", help="Template actif JSON", default="templates/bilan_actif.json")
    parser.add_argument("--tpl_pass", help="Template passif JSON", default="templates/bilan_passif.json")
    parser.add_argument("--tpl_cpc", help="Template cpc JSON", default="templates/bilan_cpc.json")
    parser.add_argument("--out", "-o", help="Chemin de sortie Excel (optionnel)", default=None)
    parser.add_argument("--use_gemini", action="store_true", help="Activer Gemini (si clés configurées)")
    args = parser.parse_args()
    if not args.pdf:
        print("Aucun PDF fourni. Exemple : python extract_with_gemini.py --pdf /chemin/file.pdf --use_gemini")
        exit(0)
    ensure_templates_exist(args.tpl_act, args.tpl_pass, args.tpl_cpc)
    out = process_pdf_with_templates(args.pdf,
                                    tpl_act_path=args.tpl_act,
                                    tpl_pass_path=args.tpl_pass,
                                    tpl_cpc_path=args.tpl_cpc,
                                    use_gemini=args.use_gemini,
                                    out_xlsx=args.out)
    print("Result:", out)

def align_numbers(nums, keys):
    L = len(keys)
    out = [''] * L
    if not nums:
        return out
    take = nums[-L:] if len(nums) >= L else nums
    start = L - len(take)
    for idx, v in enumerate(take):
        out[start + idx] = v
    return out

