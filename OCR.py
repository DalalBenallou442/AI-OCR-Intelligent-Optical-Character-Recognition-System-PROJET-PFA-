import re
import unicodedata
import pandas as pd
from collections import OrderedDict
from tabula import read_pdf
import PyPDF2

pdf_path = "test.pdf"
pages = "all"
DEBUG = False

ignore_keywords = [
    "modèle", "identification", "exercice comptable", "identifiant fiscal",
    "raison sociale", "ice", "adresse", "ville", "activité", "art. taxe",
    "exercice", "au :", "identifiant", "raison", "activité :", "modèle :"
]

number_re = re.compile(r"""[-+]?    # signe optionnel
                          \d{1,3}(?:[ \.\,]\d{3})*(?:[\,\.]\d+)?""", re.VERBOSE)

def remove_accents(s):
    if s is None: return ""
    return ''.join(ch for ch in unicodedata.normalize('NFD', str(s)) if unicodedata.category(ch) != 'Mn')

def parse_number(s):
    if s is None: return 0.0
    s = str(s).strip()
    if s == "" or s.lower() in ["nan", "none"]: return 0.0
    s = s.replace('\xa0', ' ').strip()
    # gérer formats avec ',' ou '.'
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

def normalize_text(s):
    if s is None: return ""
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def is_cpc_heading(text):
    if not text: return False
    t = remove_accents(text).upper()
    return ("COMPTE DE PRODUITS" in t) or ("COMPTE DE PRODUITS ET CHARGES" in t) or ("PRODUITS D'EXPLOITATION" in t) or ("CHARGES D'EXPLOITATION" in t) or ("PRODUITS FINANCIERS" in t) or ("PRODUITS NON COURANTS" in t) or ("CHARGES NON COURANTES" in t) or ("RESULTAT" in t and ("EXPLOITATION" in t or "COURANT" in t))

def detect_cpc_parent(cat_text):
    """Retourne Parent_Sous_Categorie standardisé si trouvable"""
    if not cat_text: return ""
    t = remove_accents(cat_text).upper()
    if "EXPLOITATION" in t:
        return "Exploitation"
    if "FINANCIER" in t:
        return "FINANCIER"
    if "NON COURANT" in t or "NON-COURANT" in t or "NONCOURANT" in t:
        return "NON COURANT"
    if "COURANT" in t:
        return "COURANT"
    # fallback: use cat_text itself trimmed
    return normalize_text(cat_text)

def canonicalize_category(cat):
    if not cat: return ""
    raw = remove_accents(cat).upper()
    raw = re.sub(r'\s+', ' ', raw).strip()

    # FORCER PROVISION -> forme standard passif
    if "PROVISION" in raw or "PROVISIONS" in raw:
        return "PROVISIONS DURABLES POUR RISQUES ET CHARGES (D)"

    # PASSIF spécifique
    if "CAPITAUX PROPRES ASSIMILES" in raw or "ASSIMILES" in raw:
        return "CAPITAUX PROPRES ASSIMILES (B)"
    if "CAPITAUX" in raw or "CAPITAUX PROPRES" in raw:
        return "CAPITAUX PROPRES"
    if "DETTES DE FINANCEMENT" in raw or "DETTE DE FINANCEMENT" in raw or "EMPRUNT" in raw or "DETT" in raw:
        return "DETTES DE FINANCEMENT (C)"
    if "ECARTS DE CONVERSION - PASSIF" in raw or "ECARTS DE CONVERSION PASSIF" in raw:
        return "ECARTS DE CONVERSION - PASSIF (E)"
    if "DETTES DU PASSIF CIRCULANT" in raw or "DETTE DU PASSIF CIRCULANT" in raw:
        return "DETTES DU PASSIF CIRCULANT (F)"
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

    # CPC / COMPTE DE PRODUITS ET CHARGES - reconnaissance simple par mots clés
    cpc_keywords = ["PRODUIT", "PRODUITS", "CHARGE", "CHARGES", "EXPLOITATION", "FINANCIER", "NON COURANT", "RESULTAT", "COMPTE DE PRODUITS"]
    for kw in cpc_keywords:
        if kw in raw:
            return raw  # renvoyer tel quel (sera utilisé pour Sous_Catégorie / Parent)

    return normalize_text(cat).upper()


def category_to_table(canonical_cat):
    if not canonical_cat:
        return "Bilan_Actif"
    c = canonical_cat.upper()

    # --- d'abord les mots-clés PASSIF (priorité) ---
    passif_keywords = ["PROVISION", "PROVISIONS", "CAPITAUX", "DETTES", "DETTE", "PASSIF", "TRESORERIE - PASSIF", "CAPITAUX PROPRES ASSIMILES", "EMPRUNT"]
    for k in passif_keywords:
        if k in c:
            return "Bilan_Passif"

    # ensuite, ACTIF
    actif_keywords = ["IMMOBILISATIONS", "STOCKS", "CREANCES", "TITRES", "TRÉSORERIE", "ÉCARTS", "ECARTS"]
    for k in actif_keywords:
        if k in c:
            return "Bilan_Actif"

    # enfin, CPC si indicateurs CPC présents
    cpc_indicators = ["PRODUIT", "PRODUITS", "CHARGE", "CHARGES", "EXPLOITATION", "FINANCIER", "NON COURANT", "RESULTAT", "COMPTE DE PRODUITS"]
    for k in cpc_indicators:
        if k in c:
            return "Bilan_CPC"

    return "Bilan_Actif"


def split_multi_line_cell(cell):
    if cell is None: return []
    s = str(cell)
    parts = re.split(r'[\n\r]+|\s{2,}|\t', s)
    parts = [p.strip() for p in parts if p]
    return parts

def extract_numeric_from_cells(cells, max_nums=4):
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
            # prend les occurrences les plus à droite
            for f in reversed(found):
                nums.insert(0, f)
                if len(nums) >= max_nums:
                    break
            # enlever le dernier nombre trouvé du contenu (pour laisser la rubrique propre)
            flat[i] = re.sub(re.escape(found[-1]) + r'\s*$', '', flat[i]).strip()
            i -= 1
        else:
            break
    nums = nums[-max_nums:]
    nums = ([""] * (max_nums - len(nums))) + nums
    left = flat[:i+1]
    rubrique = " ".join([t for t in left if t]).strip()
    return rubrique, nums

def pages_to_list(pages_str):
    if isinstance(pages_str, int):
        return [pages_str]
    if pages_str in ("all", "1-end"):
        return None
    pages_list = []
    for part in pages_str.split(','):
        part = part.strip()
        if '-' in part:
            a,b = part.split('-')
            pages_list.extend(range(int(a), int(b)+1))
        else:
            pages_list.append(int(part))
    return pages_list

def process_pdf(pdf_path, out="bilan_actif_passif_cpc_corrige.xlsx"):
    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        total_pages = len(reader.pages)

    pages_list = pages_to_list("all")
    if pages_list is None:
        pages_list = range(1, total_pages + 1)
    else:
        pages_list = [p for p in pages_list if p <= total_pages]

    data_actif = OrderedDict()
    data_passif = OrderedDict()
    data_cpc = OrderedDict()

    current_cat = ""
    current_table = "Bilan_Actif"
    last_keys = {"Bilan_Actif": None, "Bilan_Passif": None, "Bilan_CPC": None}
    last_cpc_parent = ""  # pour mémoriser Explotation/FINANCIER/NON COURANT/COURANT

    for p in pages_list:
        try:
            dfs = read_pdf(pdf_path, pages=p, multiple_tables=True, lattice=True)
        except Exception as e:
            print(f"Erreur extraction page {p}: {e}")
            continue
        if not dfs:
            continue
        if DEBUG: print(f"Page {p} -> {len(dfs)} table(s)")
        for df in dfs:
            df = df.fillna("").astype(str)
            if DEBUG:
                print("---- RAW head ----")
                print(df.head(20).to_string())
            for _, row in df.iterrows():
                cells = [c.strip() for c in row.tolist() if str(c).strip() != ""]
                joined_lower = " ".join(row.tolist()).lower()
                if any(k in joined_lower for k in ignore_keywords):
                    continue

                rubrique_raw, nums = extract_numeric_from_cells(row.tolist(), max_nums=4)

                segs = split_multi_line_cell(rubrique_raw)
                detected_cat = None
                actual_rubrique = rubrique_raw.strip() if rubrique_raw else ""

                # détecter un titre en MAJ (ex: "I. PRODUITS D'EXPLOITATION")
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

                # autre heuristique pour titres
                if not detected_cat and actual_rubrique:
                    m = re.match(r'^\s*([A-ZÀ-ÖØ-Ý0-9 \-\'\(\)\/\.]+?)\s{1,}(.+)$', actual_rubrique)
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

                # si on détecte un titre CPC explicite (ex: COMPTE DE PRODUITS ET CHARGES)
                if detected_cat and is_cpc_heading(detected_cat):
                    current_cat = canonicalize_category(detected_cat)
                    current_table = "Bilan_CPC"
                    last_cpc_parent = detect_cpc_parent(detected_cat) or last_cpc_parent
                    if DEBUG: print(f"--- CPC heading detected: {detected_cat} -> parent {last_cpc_parent}")
                elif detected_cat:
                    # mise à jour catégorie générale (actif/passif ou section CPC)
                    canon = canonicalize_category(detected_cat)
                    current_cat = canon
                    table_from_cat = category_to_table(canon)
                    if table_from_cat is not None:
                        current_table = table_from_cat
                    # si la catégorie est CPC-type, mémoriser parent si possible
                    if current_table == "Bilan_CPC":
                        last_cpc_parent = detect_cpc_parent(detected_cat) or last_cpc_parent
                    if DEBUG: print(f"Detected category: {detected_cat} -> {canon} -> table {current_table}")

                # Si ligne contient un mot-clé CPC (ex: "PRODUITS D'EXPLOITATION") même sans detected_cat
                if current_table != "Bilan_CPC" and is_cpc_heading(" ".join(row.tolist())):
                    current_table = "Bilan_CPC"
                    # essayer détecter parent
                    possible = " ".join(segs)
                    last_cpc_parent = detect_cpc_parent(possible) or last_cpc_parent
                    if DEBUG: print("Forced CPC by row content; parent:", last_cpc_parent)

                # ignorer lignes "TOTAL" quand elles n'ont pas de montants
                if actual_rubrique and re.search(r'\bTOTAL\b', actual_rubrique, flags=re.I) and all(not n for n in nums):
                    if DEBUG: print("Skip total-only line:", actual_rubrique)
                    continue

                # Convertir les 4 valeurs numériques (ordre conforme à ton PDF)
                n_prop = parse_number(nums[-4]) if len(nums) >= 4 else 0.0  # propres à l'exercice (1)
                n_prev = parse_number(nums[-3]) if len(nums) >= 3 else 0.0  # concernant exercices précédents (2)
                n_tot  = parse_number(nums[-2]) if len(nums) >= 2 else 0.0  # totaux de l'exercice (3)
                n_prevtot = parse_number(nums[-1]) if len(nums) >= 1 else 0.0  # exercice précédent (4)

                sous_cat_norm = current_cat if current_cat else ""
                rubrique_norm = normalize_text(actual_rubrique)

                target_data = data_actif if current_table == "Bilan_Actif" else (data_passif if current_table == "Bilan_Passif" else data_cpc)
                lk = last_keys.get(current_table, None)

                # ligne sans rubrique mais avec montants => ajouter aux derniers
                if (not rubrique_norm) and any(n for n in [n_prop, n_prev, n_tot, n_prevtot]):
                    if lk is not None and lk in target_data:
                        e = target_data[lk]
                        if current_table == "Bilan_CPC":
                            e["Propres a l'exercice"] += n_prop
                            e["Concernant les exercices_prec"] += n_prev
                            e["Totaux de l'exercice"] += n_tot
                            e["Exercice_prec"] += n_prevtot
                        else:
                            e["Montant_Brut"] += n_prop
                            e["Amortissements_Provisions"] += n_prev
                            e["Net_Exercice"] += n_tot
                            e["Net_Exercice_Prec"] += n_prevtot
                        continue
                    else:
                        # créer entrée agrégée vide
                        if current_table == "Bilan_CPC":
                            key = (last_cpc_parent or sous_cat_norm, "")
                            if key not in target_data:
                                target_data[key] = {
                                    "Type_Tableau": "Bilan_CPC",
                                    "Parent_Sous_Categorie": last_cpc_parent or sous_cat_norm,
                                    "Sous_Categorie": sous_cat_norm,
                                    "Rubrique": "",
                                    "Propres a l'exercice": 0.0,
                                    "Concernant les exercices_prec": 0.0,
                                    "Totaux de l'exercice": 0.0,
                                    "Exercice_prec": 0.0,
                                    "Commentaires": ""
                                }
                            e = target_data[key]
                            e["Propres a l'exercice"] += n_prop
                            e["Concernant les exercices_prec"] += n_prev
                            e["Totaux de l'exercice"] += n_tot
                            e["Exercice_prec"] += n_prevtot
                            last_keys[current_table] = key
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
                            e["Montant_Brut"] += n_prop
                            e["Amortissements_Provisions"] += n_prev
                            e["Net_Exercice"] += n_tot
                            e["Net_Exercice_Prec"] += n_prevtot
                            last_keys[current_table] = key
                            continue

                # création / mise à jour normale d'une entrée
                key = (last_cpc_parent or sous_cat_norm, rubrique_norm) if current_table == "Bilan_CPC" else (sous_cat_norm, rubrique_norm)
                if key in target_data:
                    e = target_data[key]
                    if current_table == "Bilan_CPC":
                        e["Propres a l'exercice"] += n_prop
                        e["Concernant les exercices_prec"] += n_prev
                        e["Totaux de l'exercice"] += n_tot
                        e["Exercice_prec"] += n_prevtot
                    else:
                        e["Montant_Brut"] += n_prop
                        e["Amortissements_Provisions"] += n_prev
                        e["Net_Exercice"] += n_tot
                        e["Net_Exercice_Prec"] += n_prevtot
                else:
                    if current_table == "Bilan_CPC":
                        parent = last_cpc_parent or sous_cat_norm
                        target_data[key] = {
                            "Type_Tableau": "Bilan_CPC",
                            "Parent_Sous_Categorie": parent,
                            "Sous_Categorie": sous_cat_norm,
                            "Rubrique": rubrique_norm,
                            "Propres a l'exercice": n_prop,
                            "Concernant les exercices_prec": n_prev,
                            "Totaux de l'exercice": n_tot,
                            "Exercice_prec": n_prevtot,
                            "Commentaires": ""
                        }
                    else:
                        target_data[key] = {
                            "Type_Tableau": current_table,
                            "Sous_Catégorie": sous_cat_norm,
                            "Rubrique": rubrique_norm,
                            "Montant_Brut": n_prop,
                            "Amortissements_Provisions": n_prev,
                            "Net_Exercice": n_tot,
                            "Net_Exercice_Prec": n_prevtot,
                            "Commentaires": ""
                        }
                last_keys[current_table] = key

    df_actif = pd.DataFrame(list(data_actif.values()))
    df_passif = pd.DataFrame(list(data_passif.values()))
    df_cpc = pd.DataFrame(list(data_cpc.values()))

    # nettoyage
    if "Sous_Catégorie" in df_actif.columns:
        df_actif["Sous_Catégorie"] = df_actif["Sous_Catégorie"].replace("0", "").astype(str).str.strip()
    if "Rubrique" in df_actif.columns:
        df_actif["Rubrique"] = df_actif["Rubrique"].str.strip()

    if "Sous_Catégorie" in df_passif.columns:
        df_passif["Sous_Catégorie"] = df_passif["Sous_Catégorie"].replace("0", "").astype(str).str.strip()
    if "Rubrique" in df_passif.columns:
        df_passif["Rubrique"] = df_passif["Rubrique"].str.strip()

    if "Sous_Categorie" in df_cpc.columns:
        # backward compatibility si nom de colonne varie
        df_cpc["Sous_Categorie"] = df_cpc.get("Sous_Categorie", df_cpc.get("Sous_Catégorie", "")).astype(str).str.strip()
    if "Rubrique" in df_cpc.columns:
        df_cpc["Rubrique"] = df_cpc["Rubrique"].str.strip()
    if "Parent_Sous_Categorie" in df_cpc.columns:
        df_cpc["Parent_Sous_Categorie"] = df_cpc["Parent_Sous_Categorie"].astype(str).str.strip()

    # renommer colonnes passif comme avant
    if "Net_Exercice" in df_passif.columns:
        df_passif = df_passif.rename(columns={
            "Net_Exercice": "Exercice",
            "Net_Exercice_Prec": "Exercice_Précedent"
        })

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_actif.to_excel(writer, sheet_name="Bilan_Actif", index=False)
        df_passif.to_excel(writer, sheet_name="Bilan_Passif", index=False)
        df_cpc.to_excel(writer, sheet_name="Bilan_CPC", index=False)

    print("✅ Fini — fichier :", out)
    print("Lignes Actif:", len(df_actif), " Lignes Passif:", len(df_passif), " Lignes CPC:", len(df_cpc))
    return out

if __name__ == "__main__":
    process_pdf(pdf_path)
