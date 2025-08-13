# ocr.py (version SANS OCR et SANS Camelot)
import os # Pour les opérations sur le système de fichiers (chemins, suppression de fichiers)
import re
import unicodedata
import pandas as pd
from collections import OrderedDict
from tabula import read_pdf
import PyPDF2

# -------------------------
# Phase 2 : Fonctions utilitaires
# =========================
# Ces fonctions permettent de nettoyer et de normaliser les données extraites.

# Supprime les accents d'une chaîne de caractères
def remove_accents(s):
 # Prend une chaîne et retire les accents (normalisation NFD + suppression des marques diacritiques)
    if s is None: return ""
    return ''.join(ch for ch in unicodedata.normalize('NFD', str(s)) if unicodedata.category(ch) != 'Mn')

# Convertit une chaîne en nombre float en gérant les formats locaux et les parenthèses pour les négatifs
def parse_number(s):
    if s is None: return 0.0
    s = str(s).strip()
    if s == "" or s.lower() in ["nan", "none"]: return 0.0
    s = s.replace('\xa0', ' ').strip()
    if s.count(',') and s.count('.'):
        if s.rfind('.') < s.rfind(','):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    else:
        if s.count(',') and not s.count('.'):
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    s = re.sub(r'[^\d\.\-\(\)]', '', s)
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return float(s)
    except:
        return 0.0
# Normalise le texte en supprimant les espaces superflus
def normalize_text(s):
    if s is None: return ""
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()
# Standardise les catégories comptables pour les classer correctement
def canonicalize_category(cat):
    if not cat: return ""
    raw = remove_accents(cat).upper()
    raw = re.sub(r'\s+', ' ', raw).strip()
    # PASSIF
    if "CAPITAUX PROPRES ASSIMILES" in raw or "ASSIMILES" in raw:
        return "CAPITAUX PROPRES ASSIMILES (B)"
    if "CAPITAUX" in raw or "CAPITAUX PROPRES" in raw:
        return "CAPITAUX PROPRES"
    if "DETTES DE FINANCEMENT" in raw or "DETTE DE FINANCEMENT" in raw or "EMPRUNT" in raw or "DETT" in raw:
        return "DETTES DE FINANCEMENT (C)"
    if "PROVISIONS DURABLES" in raw or "PROVISIONS POUR RISQUES" in raw or "PROVISION" in raw:
        return "PROVISIONS DURABLES POUR RISQUES ET CHARGES (D)"
    if "ECARTS DE CONVERSION - PASSIF" in raw or "ECARTS DE CONVERSION PASSIF" in raw:
        return "ECARTS DE CONVERSION - PASSIF (E)"
    if "DETTES DU PASSIF CIRCULANT" in raw or "DETTE DU PASSIF CIRCULANT" in raw:
        return "DETTES DU PASSIF CIRCULANT (F)"
    if "AUTRES PROVISIONS" in raw or "AUTRES PROVISIONS POUR RISQUES" in raw:
        return "AUTRES PROVISIONS POUR RISQUES ET CHARGES (G)"
    if "TRESORERIE - PASSIF" in raw or "TRESORERIE PASSIF" in raw or "BANQUES (SOLDES CREDITEURS)" in raw:
        return "TRESORERIE - PASSIF"
    if "TOTAL GENERAL" in raw or "TOTAL I+II+III" in raw:
        return "TOTAL GENERAL"
    # ACTIF
    if "IMMOBILISATION" in raw and ("NON" in raw or "NON VALEUR" in raw or "NONVALEUR" in raw):
        return "IMMOBILISATIONS EN NON VALEUR"
    if "IMMOBILISATION" in raw and "INCORPOREL" in raw:
        return "IMMOBILISATIONS INCORPORELLES"
    if "IMMOBILISATION" in raw and "CORPOREL" in raw:
        return "IMMOBILISATIONS CORPORELLES"
    if "IMMOBILISATION" in raw and "FINANCI" in raw:
        return "IMMOBILISATIONS FINANCIÈRES"
    if "ECART" in raw and "ACTIF" in raw:
        return "ÉCARTS DE CONVERSION – ACTIF"
    if "STOCK" in raw:
        return "STOCKS"
    if "CREANCE" in raw or "CREANCES" in raw or "ACTIF CIRCULANT" in raw:
        return "CRÉANCES ACTIF CIRCULANT"
    if "TITRE" in raw or "VALEUR" in raw or "PLACEMENT" in raw:
        return "TITRES ET VALEURS DE PLACEMENT"
    if "TRESOR" in raw or "TRÉSOR" in raw or "TRESORERIE" in raw:
        return "TRÉSORERIE ACTIF"
    if "TOTAL" in raw:
        return "TOTAL"
    return normalize_text(cat).upper()
# Détermine si la catégorie appartient au bilan actif ou passif
def category_to_table(canonical_cat):
    if not canonical_cat:
        return "Bilan_Actif"
    c = canonical_cat.upper()
    passif_keywords = ["CAPITAUX", "DETTES", "PROVISION", "PASSIF", "TRESORERIE - PASSIF", "CAPITAUX PROPRES ASSIMILES"]
    for k in passif_keywords:
        if k in c:
            return "Bilan_Passif"
    actif_keywords = ["IMMOBILISATIONS", "STOCKS", "CREANCES", "TITRES", "TRÉSORERIE", "ÉCARTS", "ECARTS"]
    for k in actif_keywords:
        if k in c:
            return "Bilan_Actif"
    return "Bilan_Actif"

# Sépare les cellules qui contiennent plusieurs lignes ou valeurs
def split_multi_line_cell(cell):
    if cell is None: return []
    s = str(cell)
    parts = re.split(r'[\n\r]+|\s{2,}|\t', s)
    parts = [p.strip() for p in parts if p]
    return parts

# =========================
# Phase 3 : Extraction depuis le PDF
# =========================
# Cette phase lit le PDF, sélectionne les pages à traiter et extrait les tableaux
# avec Tabula en mode lattice (grille) ou stream (texte)
def process_pdf(pdf_path, pages="all", debug=False, result_folder="result", out_name=None,
                remove_input=False):
    
    # Vérification de l'existence du fichier PDF et création du dossier de sortie
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"Fichier non trouvé: {pdf_path}")
    os.makedirs(result_folder, exist_ok=True)

    # Définition du nom du fichier Excel de sortie
    if out_name is None:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        out_name = f"{base}_bilan.xlsx"
    out_path = os.path.join(result_folder, out_name)

    # regex nombres local
    number_re = re.compile(r"""[-+]?    # signe optionnel
                              \d{1,3}(?:[ \.\,]\d{3})*(?:[\,\.]\d+)?""", re.VERBOSE)

    ignore_keywords = [
        "modèle", "identification", "exercice comptable", "identifiant fiscal",
        "raison sociale", "ice", "adresse", "ville", "activité", "art. taxe",
        "exercice", "au :", "identifiant", "raison", "activité :", "modèle :"
    ]

     # Lecture du PDF et nombre total de pages
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        total_pages = len(reader.pages)

    def pages_to_list_local(pages_str):
        if isinstance(pages_str, int):
            return [pages_str]
        if pages_str in ("all", "1-end"):
            return None
        pages_list_local = []
        for part in pages_str.split(','):
            part = part.strip()
            if '-' in part:
                a,b = part.split('-')
                pages_list_local.extend(range(int(a), int(b)+1))
            else:
                pages_list_local.append(int(part))
        return pages_list_local

    pages_list = pages_to_list_local(pages)
    if pages_list is None:
        pages_list = range(1, total_pages + 1)
    else:
        pages_list = [p for p in pages_list if p <= total_pages]

    data_actif = OrderedDict()
    data_passif = OrderedDict()
    current_cat = ""
    current_table = "Bilan_Actif"
    last_key = None

    def extract_numeric_from_cells_local(cells, max_nums=4):
        flat = []
        for c in cells:
            parts = split_multi_line_cell(c)
            if parts:
                flat.extend(parts)
            else:
                flat.append("")
        nums = []
        i = len(flat) - 1
        while i >= 0 and len(nums) < max_nums:
            cell = flat[i]
            if cell == "":
                i -= 1
                continue
            found = number_re.findall(cell)
            if found:
                for f in reversed(found):
                    nums.insert(0, f)
                    if len(nums) >= max_nums:
                        break
                flat[i] = re.sub(re.escape(found[-1]) + r'\s*$', '', flat[i]).strip()
                i -= 1
            else:
                break
        nums = nums[-max_nums:]
        nums = ([""] * (max_nums - len(nums))) + nums
        left = flat[:i+1]
        rubrique = " ".join([t for t in left if t]).strip()
        return rubrique, nums

    # boucle principale : pages -> tableaux -> lignes
    for p in pages_list:
        if debug: print("Traitement page:", p)
        dfs = []
        # Tabula lattice (préféré)
        try:
            dfs = read_pdf(pdf_path, pages=p, multiple_tables=True, lattice=True)
            if debug: print(f"Tabula lattice trouvé {len(dfs)} table(s) page {p}")
        except Exception as e:
            if debug: print("Tabula lattice failed:", e)

        # Si vide, tenter Tabula stream
        if not dfs:
            try:
                dfs = read_pdf(pdf_path, pages=p, multiple_tables=True, lattice=False)
                if debug: print(f"Tabula stream trouvé {len(dfs)} table(s) page {p}")
            except Exception as e:
                if debug: print("Tabula stream failed:", e)

        if not dfs:
            if debug: print("Aucun tableau trouvé page", p)
            continue

        for df in dfs:
            try:
                df = df.fillna("").astype(str)
            except Exception:
                pass
            if debug:
                print("---- RAW head ----")
                try:
                    print(df.head(10).to_string())
                except:
                    pass

            for _, row in df.iterrows():
                cells = [c.strip() for c in row.tolist()]
                joined_lower = " ".join(cells).lower()
                if any(k in joined_lower for k in ignore_keywords):
                    continue

                rubrique_raw, nums = extract_numeric_from_cells_local(cells, max_nums=4)
                segs = split_multi_line_cell(rubrique_raw)
                detected_cat = None
                actual_rubrique = rubrique_raw.strip() if rubrique_raw else ""

                # heuristiques détection catégorie
                if segs and len(segs) > 1:
                    for s in segs:
                        if s and s.upper() == s and len(re.findall(r'[A-Z]', s)) >= 3:
                            detected_cat = s.strip()
                            other = [x for x in segs if x != s]
                            actual_rubrique = " ".join(other).strip()
                            break
                    if not detected_cat and segs[0].upper() == segs[0]:
                        detected_cat = segs[0].strip()
                        actual_rubrique = " ".join(segs[1:]).strip()

                if not detected_cat and actual_rubrique:
                    m = re.match(r'^\s*([A-ZÀ-ÖØ-Ý0-9 \-\'\(\)\/]+?)\s{1,}(.+)$', actual_rubrique)
                    if m:
                        candidate = m.group(1).strip()
                        rest = m.group(2).strip()
                        if len(re.sub(r'[^A-ZÀ-ÖØ-Ý]', '', candidate)) >= 3:
                            detected_cat = candidate
                            actual_rubrique = rest
                    else:
                        maj_segments = re.findall(r'([A-ZÀ-ÖØ-Ý][A-Z0-9 \-\'\(\)\/]{2,})', actual_rubrique)
                        if maj_segments:
                            candidate = max(maj_segments, key=len).strip()
                            if len(re.sub(r'[^A-ZÀ-ÖØ-Ý]', '', candidate)) >= 3:
                                actual_rubrique = re.sub(re.escape(candidate), '', actual_rubrique, count=1).strip()
                                detected_cat = candidate

                if detected_cat:
                    canon = canonicalize_category(normalize_text(detected_cat))
                    current_cat = canon
                    table_from_cat = category_to_table(canon)
                    if table_from_cat is not None:
                        current_table = table_from_cat
                    if debug: print(f"Detected category: {detected_cat} -> {canon} -> table {current_table}")

                if actual_rubrique and re.search(r'\bTOTAL\b', actual_rubrique, flags=re.I) and all(not n for n in nums):
                    if debug: print("Skip total-only line:", actual_rubrique)
                    continue

                montant_brut = parse_number(nums[-4]) if len(nums) >= 4 else 0.0
                amort_prov   = parse_number(nums[-3]) if len(nums) >= 3 else 0.0
                net_ex       = parse_number(nums[-2]) if len(nums) >= 2 else 0.0
                net_prec     = parse_number(nums[-1]) if len(nums) >= 1 else 0.0

                sous_cat_norm = current_cat if current_cat else ""
                rubrique_norm = normalize_text(actual_rubrique)

                target_data = data_actif if current_table == "Bilan_Actif" else data_passif

                if (not rubrique_norm) and any(n for n in nums):
                    if last_key is not None and last_key in target_data:
                        e = target_data[last_key]
                        e["Montant_Brut"] += montant_brut
                        e["Amortissements_Provisions"] += amort_prov
                        e["Net_Exercice"] += net_ex
                        e["Net_Exercice_Prec"] += net_prec
                        continue
                    else:
                        key = (sous_cat_norm, "")
                        if key not in target_data:
                            target_data[key] = {
                                "Type_Tableau": current_table,
                                "Sous_Catégorie": sous_cat_norm,
                                "Rubrique": "",
                                "Montant_Brut": 0.0,
                                "Amortissements_Provisions": 0.0,
                                "Net_Exercice": 0.0,
                                "Net_Exercice_Prec": 0.0,
                                "Commentaires": ""
                            }
                        e = target_data[key]
                        e["Montant_Brut"] += montant_brut
                        e["Amortissements_Provisions"] += amort_prov
                        e["Net_Exercice"] += net_ex
                        e["Net_Exercice_Prec"] += net_prec
                        last_key = key
                        continue

                key = (sous_cat_norm, rubrique_norm)
                if key in target_data:
                    e = target_data[key]
                    e["Montant_Brut"] += montant_brut
                    e["Amortissements_Provisions"] += amort_prov
                    e["Net_Exercice"] += net_ex
                    e["Net_Exercice_Prec"] += net_prec
                else:
                    target_data[key] = {
                        "Type_Tableau": current_table,
                        "Sous_Catégorie": sous_cat_norm,
                        "Rubrique": rubrique_norm,
                        "Montant_Brut": montant_brut,
                        "Amortissements_Provisions": amort_prov,
                        "Net_Exercice": net_ex,
                        "Net_Exercice_Prec": net_prec,
                        "Commentaires": ""
                    }
                last_key = key

    # construction DataFrames finaux
    df_actif = pd.DataFrame(list(data_actif.values()))
    df_passif = pd.DataFrame(list(data_passif.values()))

    # nettoyage
    if not df_actif.empty:
        df_actif["Sous_Catégorie"] = df_actif["Sous_Catégorie"].replace("0", "").astype(str).str.strip()
        df_actif["Rubrique"] = df_actif["Rubrique"].str.strip()
    if not df_passif.empty:
        df_passif["Sous_Catégorie"] = df_passif["Sous_Catégorie"].replace("0", "").astype(str).str.strip()
        df_passif["Rubrique"] = df_passif["Rubrique"].str.strip()
        df_passif = df_passif.rename(columns={
            "Net_Exercice": "Exercice",
            "Net_Exercice_Prec": "Exercice_Précedent"
        })

    # write excel
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        if not df_actif.empty:
            df_actif.to_excel(writer, sheet_name="Actif", index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name="Actif", index=False)
        if not df_passif.empty:
            df_passif.to_excel(writer, sheet_name="Passif", index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name="Passif", index=False)

    if remove_input:
        try:
            os.remove(pdf_path)
        except:
            pass

    if debug:
        print("✅ Fini — fichier :", out_path)
        print("Lignes Actif:", 0 if df_actif.empty else len(df_actif),
              " Lignes Passif:", 0 if df_passif.empty else len(df_passif))

    return out_path
