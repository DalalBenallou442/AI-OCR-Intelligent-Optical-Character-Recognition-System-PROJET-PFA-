# -*- coding: utf-8 -*-
"""
extract_with_gemini.py (updated)
- Génération auto templates
- Robust Excel write (absolute path, try/except, fallback CSV)
- Retourne chemin fichier existant ou None
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

PROMPT_BILAN_TABLEAU = """
Tu es un assistant OCR intelligent. Tu reçois en entrée une image (ou le texte d'une SEULE page)
contenant un tableau scanné de bilan comptable en français.

Ignore tout texte avant le tableau.
Renvoie **toujours** une liste JSON (liste de dictionnaires). Pour chaque ligne fournis ces champs EXACTS si possible :
Rubrique, Montant_Brut, Amortissements_Provisions, Net_Exercice, Net_Exercice_Prec, Commentaires

Les montants peuvent être au format "12 345,67" ou "(1 234,56)". Si illisible, renvoie "".
Ne renvoie **aucun** texte hors du JSON.
"""
PROMPT_PAGE_ONLY = PROMPT_BILAN_TABLEAU

_re_number = re.compile(r'[-]?\(?\d{1,3}(?:[ \u00A0\d]{0,}\d)?[.,]\d{1,2}\)?|-?\d+')

def normalize_number_str(s):
    if s is None:
        return ""
    s = str(s).strip()
    if s in ["", "-", "—", "–", "--", "nan", "None"]:
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

# ----------------- ALL_RUBRIQUES -----------------
ALL_RUBRIQUES = {
    "Bilan_Actif": {
        "IMMOBILISATIONS EN NON VALEUR (A)": ["Immobilisations en non valeurs (A)", "Frais préliminaires", "Charges à répartir sur plusieurs exercices", "Primes de remboursement des obligations"],
        "IMMOBILISATIONS INCORPORELLES (B)": ["Immobilisations incorporelles (B)", "Immobilisations en recherche et développement", "Brevets, marques, droits et valeurs similaires", "Fonds commercial", "Autres immobilisations incorporelles"],
        "IMMOBILISATIONS CORPORELLES (C)": ["Immobilisations corporelles (C)", "Terrains", "Constructions", "Installations techniques, matériel et outillage", "Matériel de transport", "Mobilier, matériel de bureau et aménagements divers", "Autres immobilisations corporelles", "Immobilisations corporelles en cours"],
        "IMMOBILISATIONS FINANCIÈRES (D)": ["Immobilisations financières (D)", "Prêts immobilisés", "Autres créances financières", "Titres de participation", "Autres titres immobilisés"],
        "ÉCARTS DE CONVERSION – ACTIF (E)": ["Ecarts de conversion actif (E)", "Diminution des créances immobilisées", "Augmentation des dettes de financement"],
        "TOTAL I (A+B+C+D+E):": ["TOTAL I (A+B+C+D+E)"],
        "STOCKS (F)": ["Stocks (F)", "Marchandises", "Matières et fournitures consommables", "Produits en cours", "Produits intermédiaires et produits résiduels", "Produits finis"],
        "CRÉANCES DE L'ACTIF CIRCULANT (G)": ["Créances de l'actif circulant (G)", "Fournisseurs débiteurs, avances et acomptes", "Clients et comptes rattachés", "État", "Comptes d’associés", "Autres débiteurs", "Comptes de régularisation - Actif"],
        "TITRES ET VALEURS DE PLACEMENT (H)": ["Titres valeurs de placement (H)"],
        "ÉCARTS DE CONVERSION – ACTIF (I)": ["Ecarts de conversion actif (1) Elémentscirculants"],
        "TOTAL II (F+G+H+D)": ["TOTAL II (F+G+H+D)"],
        "TRÉSORERIE – ACTIF (J)": ["Trésorerie-Actif", "Banques, T.G. et C.C.P.", "Chèques et valeurs à encaisser", "Caisse, régie d’avances et accréditifs"],
        "TOTAL III ": ["TOTAL III"],
        "TOTAL GENERAL I+II+III": ["TOTAL GENERAL I+II+III"]
    },
    "Bilan_Passif": {
        "CAPITAUX PROPRES (A)": ["Capital social ou personnel (1)", "Moins : actionnaires, capital souscritnon appelé", "Moins : Capital appelé", "Moins : Dont versé", "Prime d'émission, de fusion,d'apport", "Ecarts de réévaluation", "Réserve légale", "Autres réserves", "Report à nouveau (2)", "Résultats nets en instance d'affectation(2)", "Résultat net de l'exercice (2)", "Total des capitaux propres (A)"],
        "CAPITAUX PROPRES ASSIMILES (B)": ["Capitaux propres assimilés (B)", "Subvention d'investissement", "Provisions réglementées"],
        "DETTES DE FINANCEMENT (C)": ["Dettes de financement (C)", "Emprunts obligataires", "Autres dettes de financement"],
        "PROVISIONS DURABLES POUR RISQUES ET CHARGES (D)": ["Provisions durables pour risques et charges(D)", "Provisions pour risques", "Provisions pour charges"],
        "ECARTS DE CONVERSION - PASSIF (E)": ["Ecarts de conversion-passif (E)", "Augmentation des créances immobilisées", "Diminution des dettes de financement"],
        "TOTAL I (A+B+C+D+E)": ["TOTAL I (A+B+C+D+E)"],
        "DETTES DU PASSIF CIRCULANT (F)": ["Dettes du passif circulant (F)", "Fournisseurs et comptes rattachés", "Clients créditeurs, avances et acomptes", "Personnel", "Organismes sociaux", "Etat", "Comptes d'associés", "Autres créanciers", "Comptes de régularisation passif"],
        "AUTRES PROVISIONS POUR RISQUES ET CHARGES (G)": ["Autres provisions pour risques et charges(G)"],
        "ECARTS DE CONVERSION - PASSIF (ELEMENTS CIRCULANTS) (H)": ["Ecarts de conversion - passif (Elémentscirculants) (H)"],
        "TOTAL II (F+G+H)": ["TOTAL II (F+G+H)"],
        "TRESORERIE - PASSIF": ["Crédits d'escompte", "Crédits de trésorerie", "Banques (soldes créditeurs)"],
        "TOTAL III": ["TOTAL III"],
        "TOTAL GENERAL I+II+III": ["TOTAL GENERAL I+II+III"]
    }
}

# ---------- parsing helpers (same as before) ----------
def safe_json_extract(text):
    if not text:
        return []
    match = re.search(r"\[.*\]", text, re.S)
    if not match:
        if DEBUG:
            print("⚠️ Aucun JSON détecté dans la réponse LLM (preview).")
        return []
    raw = match.group(0)
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except Exception as e:
        if DEBUG:
            print("⚠️ JSON non valide après extraction :", e)
            print("Snippet JSON:", raw[:500])
        return []

# (parse_markdown_table, parse_line_blocks, parse_gemini_text_to_rows remain identical to previous script)
# For brevity in this message I keep their definitions unchanged — assume same as earlier version.

# --- we'll re-declare them quickly (same logic) ---
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

# -------------- template helpers + ensure ------------------
def normalize_str_for_match(s):
    if s is None: return ""
    s = str(s).lower().strip()
    s = re.sub(r'[\(\)\[\]\.,;:–—\-\/\u00A0]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def best_match_name(name, candidates, cutoff=0.60):
    if not name or not candidates:
        return None
    n = normalize_str_for_match(name)
    cand_norm = {normalize_str_for_match(c): c for c in candidates}
    if n in cand_norm:
        return cand_norm[n]
    for k,v in cand_norm.items():
        if n in k or k in n:
            return v
    matches = difflib.get_close_matches(n, list(cand_norm.keys()), n=1, cutoff=cutoff)
    if matches:
        return cand_norm[matches[0]]
    return None

def load_template_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        arr = json.load(f)
    if not arr:
        return arr, []
    sample = arr[0]
    numeric_keys = [k for k in sample.keys() if k not in ("Type_Tableau","Sous_Categorie","Rubrique","Commentaires")]
    return arr, numeric_keys

def ensure_templates_exist(tpl_act_path="templates/bilan_actif.json", tpl_pass_path="templates/bilan_passif.json"):
    os.makedirs(os.path.dirname(tpl_act_path) or ".", exist_ok=True)
    os.makedirs(os.path.dirname(tpl_pass_path) or ".", exist_ok=True)
    def build_template_list(type_key):
        out = []
        rubriques_map = ALL_RUBRIQUES.get(type_key, {})
        for sous_cat, rubs in rubriques_map.items():
            if not rubs:
                rubs = ["TOTAL"]
            for rub in rubs:
                row = {
                    "Type_Tableau": type_key,
                    "Sous_Categorie": sous_cat,
                    "Rubrique": rub,
                    "Montant_Brut": "",
                    "Amortissements_Provisions": "",
                    "Net_Exercice": "",
                    "Net_Exercice_Prec": "",
                    "Commentaires": ""
                }
                out.append(row)
        return out
    if not os.path.exists(tpl_act_path):
        if DEBUG: print(f"⚠️ Template actif manquant -> création automatique : {tpl_act_path}")
        tpl_act_list = build_template_list("Bilan_Actif")
        with open(tpl_act_path, "w", encoding="utf-8") as f:
            json.dump(tpl_act_list, f, ensure_ascii=False, indent=2)
    if not os.path.exists(tpl_pass_path):
        if DEBUG: print(f"⚠️ Template passif manquant -> création automatique : {tpl_pass_path}")
        tpl_pass_list = build_template_list("Bilan_Passif")
        with open(tpl_pass_path, "w", encoding="utf-8") as f:
            json.dump(tpl_pass_list, f, ensure_ascii=False, indent=2)
    # validate quickly
    with open(tpl_act_path, 'r', encoding='utf-8') as f: json.load(f)
    with open(tpl_pass_path, 'r', encoding='utf-8') as f: json.load(f)

# -------------- call_gemini (tolérant) -------------------
def call_gemini(content, model_name="gemini-1.5-flash"):
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

# -------------- Main processing (template-first) ----------------
def process_pdf_with_templates(pdf_path, tpl_act_path, tpl_pass_path,
                               use_gemini=True, reverse_numbers=False, out_xlsx=None, fill_with_zero=False, zoom=2):
    ensure_templates_exist(tpl_act_path, tpl_pass_path)

    tpl_act, num_keys_act = load_template_json(tpl_act_path)
    tpl_pass, num_keys_pass = load_template_json(tpl_pass_path)
    candidats_act = [d.get("Rubrique","") for d in tpl_act]
    candidats_pass = [d.get("Rubrique","") for d in tpl_pass]

    if DEBUG:
        print("Templates chargés -> Actif:", len(tpl_act), "Passif:", len(tpl_pass))

    doc = fitz.open(pdf_path)
    patched_act = []
    patched_pass = []

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
                raw_text = call_gemini(payload)
            except Exception as e:
                if DEBUG: print("Gemini call exception:", e)
                raw_text = ""
        else:
            raw_text = ""

        rows = []
        if raw_text:
            rows = safe_json_extract(raw_text)
            if not rows:
                rows = parse_gemini_text_to_rows(raw_text)
        else:
            if page_text and len(page_text.strip()) > 50:
                if DEBUG: print("Gemini absent -> parsing local du texte de la page.")
                rows = parse_gemini_text_to_rows(page_text)
            else:
                rows = []

        if not rows:
            if DEBUG: print("Aucun résultat exploitable sur cette page.")
            continue

        for r in rows:
            lib = str(r.get("Rubrique","")).strip()
            nums = []
            for v in r.values():
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
            if not nums_norm and "total" not in lib.lower():
                if DEBUG: print("Skip ligne sans nombres :", lib[:80])
                continue

            ptxt = (page_text + " " + lib).lower()
            tt = None
            if "passif" in ptxt and "actif" not in ptxt:
                tt = "Bilan_Passif"
            elif "actif" in ptxt and "passif" not in ptxt:
                tt = "Bilan_Actif"

            matched = None
            if tt == "Bilan_Passif":
                matched = best_match_name(lib, candidats_pass, cutoff=0.55)
            elif tt == "Bilan_Actif":
                matched = best_match_name(lib, candidats_act, cutoff=0.55)
            else:
                matched = best_match_name(lib, candidats_act, cutoff=0.55) or best_match_name(lib, candidats_pass, cutoff=0.55)

            if not matched:
                if DEBUG: print("Aucun match template pour:", lib[:120], " -> IGNORE")
                continue

            if matched in candidats_pass:
                tpl_list = tpl_pass; nkeys = num_keys_pass; side = "passif"
            else:
                tpl_list = tpl_act; nkeys = num_keys_act; side = "actif"

            tpl_entries = [t for t in tpl_list if normalize_str_for_match(t.get("Rubrique","")) == normalize_str_for_match(matched)]
            if not tpl_entries:
                tpl_entries = [t for t in tpl_list if t.get("Rubrique","") == matched]
            if not tpl_entries:
                if DEBUG: print("Template entry introuvable malgré match -> skip")
                continue

            if reverse_numbers:
                nums_norm = list(reversed(nums_norm))

            newrow = tpl_entries[0].copy()
            for i, nk in enumerate(nkeys):
                if i < len(nums_norm):
                    val = nums_norm[i]
                    newrow[nk] = val
            newrow["_matched_from_page"] = p_index+1
            newrow["_matched_name_raw"] = lib
            newrow["_source"] = "gemini_page"
            if side == "actif":
                patched_act.append(newrow)
            else:
                patched_pass.append(newrow)

    doc.close()

    def merge_template_with_patches(template_list, patched_rows, numeric_keys):
        out = []
        patch_map = {}
        for p in patched_rows:
            key = normalize_str_for_match(p.get("Rubrique",""))
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
                out.append(row)
            else:
                out.append(t.copy())
        return out

    merged_act = merge_template_with_patches(tpl_act, patched_act, num_keys_act)
    merged_pass = merge_template_with_patches(tpl_pass, patched_pass, num_keys_pass)

    df_act = pd.DataFrame(merged_act) if merged_act else pd.DataFrame(tpl_act)
    df_pass = pd.DataFrame(merged_pass) if merged_pass else pd.DataFrame(tpl_pass)

    # AJOUT : Nettoyer toutes les valeurs NA
    df_act = df_act.fillna("")
    df_pass = df_pass.fillna("")

    # Normalize numeric strings
    for df in (df_act, df_pass):
        if df is None or df.empty:
            continue
        for col in df.columns:
            if col in ("Type_Tableau","Sous_Categorie","Rubrique","Commentaires","_matched_from_page","_matched_name_raw","_source"):
                continue
            try:
                df[col] = df[col].apply(lambda x: normalize_number_str(x) if x not in (None,"") else x)
            except Exception:
                pass

    # Try normalize sous categories
    try:
        if not df_act.empty:
            df_act["Type_Tableau"] = df_act.get("Type_Tableau", "Bilan_Actif")
            if "Sous_Categorie" in df_act.columns:
                df_act = normalize_sous_categories(df_act)
        if not df_pass.empty:
            df_pass["Type_Tableau"] = df_pass.get("Type_Tableau", "Bilan_Passif")
            if "Sous_Categorie" in df_pass.columns:
                df_pass = normalize_sous_categories(df_pass)
    except Exception as e:
        if DEBUG: print("Warning normalisation Sous_Categorie:", e)

    try:
        # Nettoyer les valeurs NA avant fill_missing_rubriques
        df_act = df_act.fillna("")
        df_pass = df_pass.fillna("")
        
        if not df_act.empty and set(["Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec"]).issubset(set(df_act.columns)):
            df_act2 = df_act[["Type_Tableau","Sous_Categorie","Rubrique","Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec","Commentaires"]]
            df_act = fill_missing_rubriques(df_act2, "Bilan_Actif", fill_with_zero=fill_with_zero)
    except Exception as e:
        if DEBUG: print("Warning fill_missing_rubriques:", e)

    # ----------------- ROBUST EXPORT -----------------
    if out_xlsx is None:
        base = Path(pdf_path).stem
        out_xlsx = f"result/{base}_patched_templates.xlsx"

    # make absolute path to avoid cwd issues
    out_xlsx = os.path.abspath(out_xlsx)
    out_dir = os.path.dirname(out_xlsx)
    os.makedirs(out_dir, exist_ok=True)

    excel_written = False
    try:
        # prefer openpyxl if available; pandas will choose engine automatically if not specified
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            if not df_act.empty:
                df_act.to_excel(writer, sheet_name="Actif", index=False)
            else:
                pd.DataFrame(tpl_act).to_excel(writer, sheet_name="Actif", index=False)
            if not df_pass.empty:
                df_pass.to_excel(writer, sheet_name="Passif", index=False)
            else:
                pd.DataFrame(tpl_pass).to_excel(writer, sheet_name="Passif", index=False)
            # flush/close inside context
        excel_written = os.path.isfile(out_xlsx)
        if excel_written and DEBUG:
            print("✅ Excel écrit:", out_xlsx)
    except Exception as e:
        # capture error and fallback to CSV
        if DEBUG:
            print("❌ Erreur écriture Excel :", repr(e))
        # fallback CSV
        try:
            csv_act = os.path.join(out_dir, Path(out_xlsx).stem + "_Actif.csv")
            csv_pass = os.path.join(out_dir, Path(out_xlsx).stem + "_Passif.csv")
            if not df_act.empty:
                df_act.to_csv(csv_act, index=False, encoding='utf-8-sig')
            else:
                pd.DataFrame(tpl_act).to_csv(csv_act, index=False, encoding='utf-8-sig')
            if not df_pass.empty:
                df_pass.to_csv(csv_pass, index=False, encoding='utf-8-sig')
            else:
                pd.DataFrame(tpl_pass).to_csv(csv_pass, index=False, encoding='utf-8-sig')
            if DEBUG:
                print("✅ Fallback CSV écrits:", csv_act, csv_pass)
            # set out_xlsx to first CSV to return something (Flask can be adjusted to return ZIP or CSV)
            out_xlsx = csv_act
            excel_written = True
        except Exception as e2:
            if DEBUG:
                print("❌ Erreur fallback CSV:", repr(e2))
            excel_written = False

    if not excel_written:
        if DEBUG:
            print("⚠️ Aucun fichier de sortie généré.")
        return None

    return out_xlsx

# -------------- normalize_sous_categories / fill_missing_rubriques (unchanged) --------------
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
        cand = difflik.get_close_matches(s0, keys_norm, n=1, cutoff=0.55)
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
    
    # CORRECTION: Utiliser "" au lieu de pd.NA
    default_val = "0.00" if fill_with_zero else ""
    
    for sous_cat in ordre_sous_categories:
        rubriques_attendues = ALL_RUBRIQUES.get(type_tableau, {}).get(sous_cat, [])
        if not rubriques_attendues and "TOTAL" in sous_cat:
            rubriques_attendues = ["TOTAL"]
        
        # CORRECTION: Vérification plus robuste pour éviter l'erreur NA
        df_sc = pd.DataFrame()
        if not df.empty:
            try:
                df_sc = df[df["Sous_Categorie"] == sous_cat]
            except Exception:
                df_sc = pd.DataFrame()
        
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
            out_df[c] = ""  # CORRECTION: "" au lieu de pd.NA
    return out_df[cols]

# ------------- wrapper utilisé par Flask ----------------
def process_inputs(pdf_path, out_xlsx=None):
    tpl_act = "templates/bilan_actif.json"
    tpl_pass = "templates/bilan_passif.json"
    ensure_templates_exist(tpl_act, tpl_pass)
    
    result = process_pdf_with_templates(
        pdf_path,
        tpl_act,
        tpl_pass,
        use_gemini=True,
        reverse_numbers=False,
        out_xlsx=out_xlsx,
        fill_with_zero=False,
        zoom=2
    )
    
    # AJOUT DEBUG
    print(f"DEBUG process_inputs: result = {result}")
    print(f"DEBUG process_inputs: result exists? {os.path.exists(result) if result else 'result is None'}")
    
    return result

# ------------- CLI (facultatif) -------------------------
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Traitement PDF -> Excel (template-first).")
    parser.add_argument("--pdf", "-p", help="Chemin vers le PDF à traiter", required=False)
    parser.add_argument("--tpl_act", help="Template actif JSON", default="templates/bilan_actif.json")
    parser.add_argument("--tpl_pass", help="Template passif JSON", default="templates/bilan_passif.json")
    parser.add_argument("--out", "-o", help="Chemin de sortie Excel (optionnel)", default=None)
    parser.add_argument("--use_gemini", action="store_true", help="Activer Gemini (si clés configurées)")
    args = parser.parse_args()
    if not args.pdf:
        print("Aucun PDF fourni. Exemple : python extract_with_gemini.py --pdf /chemin/file.pdf --use_gemini")
        exit(0)
    ensure_templates_exist(args.tpl_act, args.tpl_pass)
    out = process_pdf_with_templates(args.pdf, args.tpl_act, args.tpl_pass, use_gemini=args.use_gemini, out_xlsx=args.out)
    print("Result:", out)
