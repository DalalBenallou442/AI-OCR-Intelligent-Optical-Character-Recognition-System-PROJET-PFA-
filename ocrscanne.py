#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ocrscanne_fixed.py
Version corrigée et améliorée de ocrscanne.py
- meilleure gestion des écritures (PermissionError / fallback)
- relance Tesseract en mode numérique pour cellules contenant des chiffres
- détection automatique des colonnes numériques si pas d'en-tête
- correction d'un bug: ajout manquant des lignes au grid_text
- nettoyage final des strings (strip) sans applymap dépréciée
- sauvegarde atomique via tmp + os.replace (Windows-friendly)

Usage:
    python ocrscanne_fixed.py --pdf "test.pdf" --out_dir "C:\\Temp\\ocr_out" --zoom 4.0

Dépendances: pytesseract, Pillow, opencv-python, numpy, pandas
Optionnel: PyMuPDF (fitz) ou pdf2image pour lire des PDF ; sklearn pour clustering
"""
import os
import re
import argparse
from pathlib import Path
from PIL import Image
import numpy as np
import cv2
import pandas as pd
import tempfile
import datetime
import uuid
import csv

# ---------- CONFIG TARGET ----------
TARGET_COLUMNS = [
    "Type_Tableau", "Sous_Catégorie", "Rubrique", "Montant_Brut",
    "Amortissements_Provisions", "Net_Exercice", "Net_Exercice_Prec", "Commentaires"
]

# Liste ordonnée voulue par l'utilisateur (modifiable si besoin)
DESIRED_ROWS = [
    ("IMMOBILISATIONS EN NON VALEUR","Frais préliminaires"),
    ("IMMOBILISATIONS EN NON VALEUR","Charges à répartir sur plusieurs exercices"),
    ("IMMOBILISATIONS EN NON VALEUR","Primes de remboursement des obligations"),
    ("IMMOBILISATIONS INCORPORELLES","Immobilisations en recherche et développement"),
    ("IMMOBILISATIONS INCORPORELLES","Brevets, marques, droits et valeurs similaires"),
    ("IMMOBILISATIONS INCORPORELLES","Fonds commercial"),
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
    ("ÉCARTS DE CONVERSION – ACTIF","(Éléments courants)"),
    ("TRÉSORERIE ACTIF","Chèques et valeurs à encaisser"),
    ("TRÉSORERIE ACTIF","Banques, Trésor et CCP"),
    ("TRÉSORERIE ACTIF","Caisse, régies d’avances et accréditifs"),
]

# ---------- dependencies optional ----------
try:
    from rapidfuzz import process as rf_process, fuzz as rf_fuzz
    USE_RAPIDFUZZ = True
except Exception:
    USE_RAPIDFUZZ = False

# ---------- helpers ----------

def ensure_dir(p):
    Path(p).mkdir(parents=True, exist_ok=True)


def render_pdf_page(path, page_number=0, zoom=3.0):
    """Render a PDF page as numpy RGB image using fitz or pdf2image.
    Falls back with clear error if neither is installed.
    """
    try:
        import fitz
        doc = fitz.open(path)
        page = doc.load_page(page_number)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return np.array(img)
    except Exception:
        try:
            from pdf2image import convert_from_path
            pages = convert_from_path(path, dpi=int(72 * zoom))
            img = pages[page_number]
            return np.array(img)
        except Exception as ex:
            raise RuntimeError("Impossible de rendre le PDF : installez PyMuPDF (fitz) ou pdf2image.") from ex


def adaptive_binarize(img_gray, blocksize=35, C=15):
    return cv2.adaptiveThreshold(img_gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                                 cv2.THRESH_BINARY_INV, blockSize=blocksize, C=C)


def morphological_clean(bin_img, small_kernel=(3,3), iterations=1):
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, small_kernel)
    return cv2.morphologyEx(bin_img, cv2.MORPH_OPEN, kernel, iterations=iterations)


def detect_hv_lines(bin_img, img_shape, horiz_scale_factor=30, vert_scale_factor=30):
    h, w = img_shape[:2]
    horiz = bin_img.copy()
    vert = bin_img.copy()
    horiz_kernel_len = max(10, w // horiz_scale_factor)
    vert_kernel_len = max(10, h // vert_scale_factor)
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (horiz_kernel_len, 1))
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vert_kernel_len))

    horiz = cv2.erode(horiz, horiz_kernel, iterations=1)
    horiz = cv2.dilate(horiz, horiz_kernel, iterations=1)

    vert = cv2.erode(vert, vert_kernel, iterations=1)
    vert = cv2.dilate(vert, vert_kernel, iterations=1)

    return horiz, vert


def find_cell_candidates(table_mask, min_area=500, min_w=20, min_h=10, max_w_ratio=0.95, max_h_ratio=0.95):
    cnts, _ = cv2.findContours(table_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    H, W = table_mask.shape
    boxes = []
    for c in cnts:
        x,y,w,h = cv2.boundingRect(c)
        area = w*h
        if (area > min_area) and (w > min_w) and (h > min_h) and (w < W*max_w_ratio) and (h < H*max_h_ratio):
            boxes.append((x,y,w,h,area))
    boxes = sorted(boxes, key=lambda b: (b[1], b[0]))
    return boxes


def ocr_cell_tesseract(img_crop, lang='fra', psm=6, oem=3):
    try:
        import pytesseract
    except Exception as e:
        raise RuntimeError("pytesseract non installé. pip install pytesseract") from e
    pil = Image.fromarray(img_crop)
    cfg = f"--oem {oem} --psm {psm}"
    txt = pytesseract.image_to_string(pil, lang=lang, config=cfg)
    txt = re.sub(r'[\r\n]+', ' ', txt)
    txt = re.sub(r'\s+', ' ', txt).strip()
    return txt

# ---------- parsing & cleaning ----------

def parse_number(s):
    """Return float or None with FR/EN heuristics (handles spaces, dots, commas, parentheses)."""
    if s is None: return None
    s = str(s).strip()
    if s == "": return None
    s = s.replace('\xa0', ' ').strip()
    s = re.sub(r'[^0-9\-\(\)\., ]', '', s)
    if s == '' or re.match(r'^[\-\(\)\s]*$', s): return None
    neg = False
    if '(' in s and ')' in s:
        neg = True
        s = s.replace('(', '').replace(')', '')
    s = s.replace(' ', '')
    if '.' in s and ',' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    else:
        if s.count(',') == 1 and s.count('.') == 0:
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    try:
        v = float(s)
        if neg:
            v = -abs(v)
        return v
    except:
        return None


def detect_numeric_columns_from_grid(df_grid, min_numeric_ratio=0.35):
    """
    Retourne la liste d'indices de colonnes considérées numériques dans df_grid.
    min_numeric_ratio : fraction minimale de cellules contenant des nombres pour
    considérer une colonne comme numérique.
    """
    cols = df_grid.columns.tolist()
    numeric_scores = []
    for c in cols:
        cells = df_grid[c].astype(str).fillna('').tolist()
        if not cells:
            numeric_scores.append(0.0)
            continue
        num_count = 0
        total = 0
        for cell in cells:
            s = str(cell).strip()
            if s == "" or s.lower().startswith("page") or len(s) < 1:
                total += 1
                continue
            # quicker heuristic : presence of digit or parentheses for negatives
            if re.search(r'[\d\(\)]', s):
                num_count += 1
            total += 1
        ratio = (num_count / max(1, total))
        numeric_scores.append(ratio)
    # candidate numeric columns where ratio >= threshold
    candidate_idxs = [i for i, r in enumerate(numeric_scores) if r >= min_numeric_ratio]
    # if none found, take top-4 highest ratios
    if not candidate_idxs:
        ordered = sorted(range(len(numeric_scores)), key=lambda i: numeric_scores[i], reverse=True)
        candidate_idxs = [i for i in ordered[:4] if numeric_scores[i] > 0]
    return candidate_idxs, numeric_scores


def is_category_line(txt):
    """Heuristic to detect category headings."""
    if not txt: return False
    s = str(txt).strip()
    if re.search(r'\d', s): 
        return False
    letters = re.findall(r'[A-Za-zÀ-ÖØ-öø-ÿ]', s)
    if not letters:
        return False
    up = sum(1 for ch in s if ch.isupper())
    ratio = up / max(1, len(letters))
    if ratio > 0.5 and len(s) < 100:
        return True
    keywords = ['IMMOBILISATION','IMMOBILISATIONS','STOCK','CREANCE','CRÉANCE','TITRES','TRÉSORERIE','BILAN','ACTIF','PASSIF']
    for k in keywords:
        if k.lower() in s.lower():
            return True
    if s.isupper() and len(s) < 100:
        return True
    return False


def merge_continuation_rows(df_grid):
    """Merge consec rows without numbers (continuation of text)."""
    rows = []
    i = 0
    while i < len(df_grid):
        row = list(df_grid.iloc[i].astype(str))
        if i+1 < len(df_grid):
            next_row = list(df_grid.iloc[i+1].astype(str))
            nums_curr = any(re.search(r'\d', c) for c in row)
            nums_next = any(re.search(r'\d', c) for c in next_row)
            if (not nums_curr) and (not nums_next) and row[0].strip() and next_row[0].strip():
                merged = row.copy()
                merged[0] = (row[0] + " " + next_row[0]).strip()
                rows.append(merged)
                i += 2
                continue
        rows.append(row)
        i += 1
    if rows:
        return pd.DataFrame(rows, columns=df_grid.columns)
    return df_grid


def collapse_and_normalize_grid(df_grid):
    """Map raw grid -> normalized final dataframe (TARGET_COLUMNS)
       Improved: detect numeric columns automatically if headers mapping fails.
    """
    # detect header (inchangé)
    header_idx = None
    for i in range(min(6, len(df_grid))):
        row_s = " ".join([str(x).lower() for x in df_grid.iloc[i].astype(str)])
        if any(k in row_s for k in ['rubrique','montant','net','amort','provision','sous','categorie','type','brut']):
            header_idx = i
            break

    if header_idx is not None:
        headers = [str(x).strip() for x in df_grid.iloc[header_idx].tolist()]
        data_rows = df_grid.iloc[header_idx+1:].reset_index(drop=True)
        df_detected = pd.DataFrame(data_rows.values, columns=headers)
    else:
        df_detected = df_grid.copy()
        # generate default col names col_0, col_1 ...
        df_detected.columns = [f"c{j}" for j in range(len(df_detected.columns))]

    # fuzzy mapping of column names to TARGET_COLUMNS (existing logic)
    detected_cols = list(df_detected.columns)
    mapping = {}
    for t in TARGET_COLUMNS:
        tl = t.lower()
        best = None; best_score = -1
        for d in detected_cols:
            dlow = str(d).lower()
            score = 0
            if 'montant' in dlow or 'brut' in dlow or 'mt' in dlow: score += 40
            if 'amort' in dlow or 'provision' in dlow: score += 40
            if 'net' in dlow: score += 30
            if 'rubrique' in dlow or 'libelle' in dlow or 'designation' in dlow: score += 40
            if 'sous' in dlow or 'categorie' in dlow or 'catégorie' in dlow: score += 30
            if 'type' in dlow: score += 20
            tok_overlap = len(set(tl.split('_')) & set(dlow.split()))
            score += tok_overlap * 10
            if score > best_score:
                best_score = score; best = d
        mapping[t] = best if best_score >= 0 else None

    # If numeric columns not mapped (or empty), detect numeric columns automatically
    numeric_candidates, numeric_scores = detect_numeric_columns_from_grid(df_detected, min_numeric_ratio=0.30)
    # Choose up to 4 numeric columns, prefer rightmost order (visual)
    numeric_candidates_sorted = sorted(numeric_candidates)
    # if header mapping gave numeric columns, keep them, otherwise assign from numeric_candidates_sorted
    mapped_numeric = []
    for colname in ['Montant_Brut','Amortissements_Provisions','Net_Exercice','Net_Exercice_Prec']:
        if mapping.get(colname) and mapping[colname] in df_detected.columns:
            mapped_numeric.append(mapping[colname])
        else:
            mapped_numeric.append(None)

    # fill unmapped numeric slots from detected numeric columns (rightmost first)
    unmapped_idxs = [i for i, m in enumerate(mapped_numeric) if m is None]
    pick = numeric_candidates_sorted[-len(unmapped_idxs):] if unmapped_idxs and numeric_candidates_sorted else []
    for i, ui in enumerate(unmapped_idxs):
        if i < len(pick):
            mapped_numeric[ui] = df_detected.columns[pick[-(i+1)]]  # rightmost first

    # apply mapping into rows
    rows_out = []
    last_sous = ""
    last_type = "Bilan_Actif"
    for idx in range(len(df_detected)):
        r = df_detected.iloc[idx]
        out = {k:"" for k in TARGET_COLUMNS}
        out['Type_Tableau'] = last_type

        # Fill rubrique & sous catégorie if mapping exists
        if mapping.get('Rubrique'):
            out['Rubrique'] = "" if pd.isna(r.get(mapping['Rubrique'], "")) else str(r.get(mapping['Rubrique'], "")).strip()
        else:
            # if no rubrique mapping, assume first column is the label
            out['Rubrique'] = str(r.iloc[0]).strip() if len(r) > 0 else ""

        if mapping.get('Sous_Catégorie'):
            out['Sous_Catégorie'] = "" if pd.isna(r.get(mapping['Sous_Catégorie'], "")) else str(r.get(mapping['Sous_Catégorie'], "")).strip()
        else:
            # leave blank to be filled from context
            out['Sous_Catégorie'] = ""

        # Fill numeric fields from mapped_numeric list
        for i, numcol_name in enumerate(['Montant_Brut','Amortissements_Provisions','Net_Exercice','Net_Exercice_Prec']):
            mapped_col = mapped_numeric[i]
            val = ""
            if mapped_col and mapped_col in df_detected.columns:
                val = r.get(mapped_col, "")
                if pd.isna(val):
                    val = ""
            else:
                # fallback: try from rightmost columns by index
                try:
                    # pick column index from the end: - (i+1)
                    val = r.iloc[-(i+1)]
                except Exception:
                    val = ""
            out[numcol_name] = "" if pd.isna(val) else str(val).strip()

        # detect if current row is a category header line and treat accordingly
        combined_text = " ".join([out.get('Rubrique',''), out.get('Sous_Catégorie','')]).strip()
        if combined_text and is_category_line(combined_text):
            out['Sous_Catégorie'] = combined_text
            out['Rubrique'] = ""
            out['Montant_Brut'] = out['Amortissements_Provisions'] = out['Net_Exercice'] = out['Net_Exercice_Prec'] = ""
        if out['Sous_Catégorie'].strip():
            last_sous = out['Sous_Catégorie'].strip()
        else:
            out['Sous_Catégorie'] = last_sous
        # parse numeric cols
        for numcol in ['Montant_Brut','Amortissements_Provisions','Net_Exercice','Net_Exercice_Prec']:
            parsed = parse_number(out.get(numcol,""))
            out[numcol] = parsed if parsed is not None else ""
        # skip empty rows
        has_text = any(str(out[c]).strip() for c in ['Sous_Catégorie','Rubrique','Commentaires'])
        has_num = any(out[c] != "" for c in ['Montant_Brut','Amortissements_Provisions','Net_Exercice','Net_Exercice_Prec'])
        if not (has_text or has_num):
            continue
        rows_out.append(out)

    final_df = pd.DataFrame(rows_out, columns=TARGET_COLUMNS)
    # strip string columns safely
    for col in final_df.select_dtypes(include=['object']).columns:
        final_df[col] = final_df[col].astype(str).str.strip()
    return final_df, mapping

# ---------- safe write helpers ----------

def is_writable_dir(path_dir: Path) -> bool:
    try:
        path_dir = Path(path_dir).resolve()
        path_dir.mkdir(parents=True, exist_ok=True)
        testfile = path_dir / f".write_test_{uuid.uuid4().hex}.tmp"
        with open(testfile, "w", encoding="utf-8") as f:
            f.write("ok")
        testfile.unlink()
        return True
    except Exception:
        return False


def safe_write_csv(df, target_path: Path, **to_csv_kwargs):
    """Write CSV atomically. Accepts the same kwargs as pandas.DataFrame.to_csv.

    Ensures we do not pass duplicate keyword arguments (previous bug: passing
    index twice caused `got multiple values for keyword argument 'index'`).
    Default: index=False when caller doesn't provide it.
    """
    target_path = Path(target_path)
    target_dir = target_path.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    tmp = target_dir / (target_path.name + f".tmp.{uuid.uuid4().hex}")
    try:
        # avoid passing index twice — if caller didn't set it, default to False
        if 'index' not in to_csv_kwargs:
            to_csv_kwargs['index'] = False
        df.to_csv(tmp, **to_csv_kwargs)
        os.replace(str(tmp), str(target_path))
        return target_path
    except Exception:
        try:
            if tmp.exists():
                tmp.unlink()
        except Exception:
            pass
        raise

# ---------- final structuring / ordering ----------

def normalize_key(s):
    if s is None:
        return ""
    ss = str(s).strip().lower()
    ss = ss.replace('é','e').replace('è','e').replace('ê','e').replace('à','a').replace('ù','u').replace('ç','c')
    ss = ss.replace("’","'").replace("´","'")
    ss = re.sub(r'\s+', ' ', ss)
    return ss


def try_to_float(x):
    if x is None or x == "": return None
    try:
        if isinstance(x, (int,float)):
            return float(x)
        s = str(x).strip().replace('\xa0','').replace(' ','').replace(',', '.')
        return float(s)
    except:
        return None


def format_number_for_csv(v):
    if v is None or v == "":
        return ""
    try:
        fv = float(v)
    except:
        return ""
    if abs(fv - int(fv)) < 0.005:
        s = f"{int(round(fv))}"
    else:
        s = f"{fv:.2f}"
    return s.replace('.', ',')

# ---------- pipeline principal ----------

def pipeline(pdf_path, page=0, zoom=3.0, out_dir="./output", visualize=False, debug_save_cells=True, psm=6):
    ensure_dir(out_dir)
    print(f"[+] Rendering page {page} from PDF '{pdf_path}' (zoom {zoom}) ...")
    img = render_pdf_page(pdf_path, page_number=page, zoom=zoom)
    orig = img.copy()
    H, W = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

    print("[+] Binarization & cleaning ...")
    bin_img = adaptive_binarize(gray, blocksize=35, C=15)
    clean = morphological_clean(bin_img, small_kernel=(3,3), iterations=1)

    print("[+] Detecting horizontal & vertical lines ...")
    horiz, vert = detect_hv_lines(clean, img.shape, horiz_scale_factor=30, vert_scale_factor=30)
    table_mask = cv2.bitwise_or(horiz, vert)

    print("[+] Finding cell candidates (contours) ...")
    boxes = find_cell_candidates(table_mask, min_area=800, min_w=30, min_h=12)
    if len(boxes) < 5:
        boxes = find_cell_candidates(table_mask, min_area=300, min_w=20, min_h=8)

    overlay = orig.copy()
    for (x,y,w_box,h_box,_) in boxes:
        cv2.rectangle(overlay, (x,y), (x+w_box,y+h_box), (255,0,0), 2)

    lines_vis = orig.copy()
    hcnts, _ = cv2.findContours(horiz, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for c in hcnts:
        x,y,w_box,h_box = cv2.boundingRect(c)
        cv2.rectangle(lines_vis, (x,y), (x+w_box,y+h_box), (0,255,0), 2)
    vcnts, _ = cv2.findContours(vert, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for c in vcnts:
        x,y,w_box,h_box = cv2.boundingRect(c)
        cv2.rectangle(lines_vis, (x,y), (x+w_box,y+h_box), (0,0,255), 2)

    cv2.imwrite(os.path.join(out_dir, "debug_overlay_cells.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
    cv2.imwrite(os.path.join(out_dir, "debug_lines_hv.png"), cv2.cvtColor(lines_vis, cv2.COLOR_RGB2BGR))
    cv2.imwrite(os.path.join(out_dir, "debug_binarized.png"), bin_img)

    print(f"[+] Candidate boxes: {len(boxes)}")
    if not boxes:
        raise RuntimeError("Aucune boîte détectée — ajuste les paramètres (zoom, kernels, min_area).")

    # group boxes into rows by y proximity
    rows = []
    current = []
    y_tol = max(8, H // 300)
    last_y = None
    for b in boxes:
        x,y,w_box,h_box,_ = b
        if not current:
            current = [b]; last_y = y
        else:
            if abs(y - last_y) <= y_tol:
                current.append(b)
                last_y = int(np.mean([bb[1] for bb in current]))
            else:
                rows.append(current)
                current = [b]
                last_y = y
    if current:
        rows.append(current)

    max_cols = max(len(r) for r in rows) if rows else 0

    # cluster x centers into column centers
    xcenters = []
    for r in rows:
        for (x,y,w_box,h_box,_) in r:
            xcenters.append(x + w_box / 2.0)

    centers = []
    if xcenters and max_cols > 0:
        try:
            from sklearn.cluster import KMeans
            k = min(max_cols, max(1, len(xcenters)))
            km = KMeans(n_clusters=k, random_state=0).fit(np.array(xcenters).reshape(-1,1))
            centers = sorted([c[0] for c in km.cluster_centers_])
        except Exception:
            centers = [ (i + 0.5) * W / max(1, max_cols) for i in range(max_cols) ]
    else:
        centers = [ (i + 0.5) * W / max(1, max_cols) for i in range(max_cols) ]

    grid_text = []
    grid_boxes = []
    cells_dir = os.path.join(out_dir, "cells")
    ensure_dir(cells_dir)

    # ensure pytesseract exists
    try:
        import pytesseract
    except Exception:
        raise RuntimeError("pytesseract non installé. pip install pytesseract")

    print("[+] OCR cell-by-cell (Tesseract)...")
    for i, row in enumerate(rows):
        row_sorted = sorted(row, key=lambda b: b[0])
        row_texts = [""] * len(centers)
        row_boxrefs = [None] * len(centers)
        for (x,y,w_box,h_box,_) in row_sorted:
            cx = x + w_box/2.0
            if centers:
                col_idx = int(np.argmin([abs(cx - c) for c in centers]))
            else:
                col_idx = 0
            pad = 2
            x1 = max(0, x - pad); y1 = max(0, y - pad)
            x2 = min(W, x + w_box + pad); y2 = min(H, y + h_box + pad)
            crop = orig[y1:y2, x1:x2]

            try:
                # --- OCR initial (texte) ---
                txt = ocr_cell_tesseract(crop, lang='fra', psm=psm)

                # Si on détecte des chiffres dans le résultat, relancer Tesseract en mode "numérique"
                if re.search(r'\d', txt):
                    try:
                        cfg_num = "--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789,.-()"
                        from pytesseract import image_to_string as tesseract_img2str
                        pil = Image.fromarray(crop)
                        txt_num = tesseract_img2str(pil, lang='fra', config=cfg_num)
                        txt_num = re.sub(r'[\r\n]+', ' ', txt_num)
                        txt_num = re.sub(r'\s+', ' ', txt_num).strip()
                        if sum(c.isdigit() for c in txt_num) >= max(2, sum(c.isdigit() for c in txt)):
                            txt = txt_num
                    except Exception:
                        pass
            except Exception:
                txt = ""

            if row_texts[col_idx]:
                row_texts[col_idx] = (row_texts[col_idx] + " " + txt).strip()
            else:
                row_texts[col_idx] = txt
            row_boxrefs[col_idx] = (x1,y1,x2-x1,y2-y1)
            if debug_save_cells and i < 400:
                fname = os.path.join(cells_dir, f"cell_r{i}_c{col_idx}.png")
                try:
                    cv2.imwrite(fname, cv2.cvtColor(crop, cv2.COLOR_RGB2BGR))
                except Exception:
                    pass
        # IMPORTANT: ajouter la ligne au grid_text
        grid_text.append(row_texts)
        grid_boxes.append(row_boxrefs)

    df_grid = pd.DataFrame(grid_text)

    # merge continuation rows
    df_grid = merge_continuation_rows(df_grid)

    # collapse & normalize grid -> preliminary final_df
    preliminary_df, mapping = collapse_and_normalize_grid(df_grid)

    # ---------- Build the strictly ordered final_structured dataframe ----------
    # prepare lookup
    for c in ["Montant_Brut","Amortissements_Provisions","Net_Exercice","Net_Exercice_Prec"]:
        if c not in preliminary_df.columns:
            preliminary_df[c] = ""
    lookup = preliminary_df.copy().astype(str).fillna('')
    lookup['__norm'] = (lookup.get('Sous_Catégorie','').fillna('') + "||" + lookup.get('Rubrique','').fillna('')).map(normalize_key)
    available_keys = lookup['__norm'].tolist()

    rows_out = []
    for (s,r) in DESIRED_ROWS:
        s_n = normalize_key(s)
        r_n = normalize_key(r)
        target_key = f"{s_n}||{r_n}"
        matched_row = None

        # exact match
        exact = lookup[lookup['__norm']==target_key]
        if not exact.empty:
            matched_row = exact.iloc[0]
        else:
            # fuzz on rubrique then composite
            if USE_RAPIDFUZZ:
                candidates = lookup['__norm'].tolist()
                best = rf_process.extractOne(r_n, candidates, scorer=rf_fuzz.partial_ratio)
                if best and best[1] >= 70:
                    matched_row = lookup[lookup['__norm']==best[0]].iloc[0]
                else:
                    best2 = rf_process.extractOne(target_key, candidates, scorer=rf_fuzz.token_sort_ratio)
                    if best2 and best2[1] >= 65:
                        matched_row = lookup[lookup['__norm']==best2[0]].iloc[0]
            else:
                # fallback substring search
                for k in available_keys:
                    if r_n in k:
                        matched_row = lookup[lookup['__norm']==k].iloc[0]
                        break
                if matched_row is None:
                    for k in available_keys:
                        if s_n in k:
                            matched_row = lookup[lookup['__norm']==k].iloc[0]
                            break

        if matched_row is not None:
            mb = try_to_float(matched_row.get('Montant_Brut',''))
            am = try_to_float(matched_row.get('Amortissements_Provisions',''))
            ne = try_to_float(matched_row.get('Net_Exercice',''))
            npv = try_to_float(matched_row.get('Net_Exercice_Prec',''))
        else:
            mb = am = ne = npv = None

        rows_out.append({
            "Type_Tableau": "Bilan_Actif",
            "Sous_Catégorie": s,
            "Rubrique": r,
            "Montant_Brut": format_number_for_csv(mb),
            "Amortissements_Provisions": format_number_for_csv(am),
            "Net_Exercice": format_number_for_csv(ne),
            "Net_Exercice_Prec": format_number_for_csv(npv),
            "Commentaires": ""
        })

    final_structured_df = pd.DataFrame(rows_out, columns=TARGET_COLUMNS)

    # Save reconstructed_table (raw) and final structured Excel (.xlsx)
    csv_raw = Path(out_dir) / "reconstructed_table.xlsx"
    final_xlsx_path = Path(out_dir) / "final_structured.xlsx"

    # prefer Excel; fallback to CSV if openpyxl not available
    try:
        import openpyxl  # noqa: F401
        EXCEL_AVAILABLE = True
    except Exception:
        EXCEL_AVAILABLE = False

    def safe_write_excel(df, target_path: Path, sheet_name: str = "Sheet1"):
        target_path = Path(target_path)
        target_dir = target_path.parent
        target_dir.mkdir(parents=True, exist_ok=True)
        tmp = target_dir / (target_path.name + f".tmp.{uuid.uuid4().hex}")
        try:
            with pd.ExcelWriter(tmp, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            os.replace(str(tmp), str(target_path))
            return target_path
        except Exception:
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass
            raise

    attempts = []
    written = False
    primary_dir = Path(out_dir).resolve()
    fallback_dirs = [Path(tempfile.gettempdir()), Path(os.path.join(os.environ.get("USERPROFILE",""), "Documents"))]

    if EXCEL_AVAILABLE:
        try:
            if is_writable_dir(primary_dir):
                safe_write_excel(preliminary_df, csv_raw, sheet_name='Reconstructed')
                safe_write_excel(final_structured_df, final_xlsx_path, sheet_name='Final')
                written = True
                attempts.append(str(primary_dir))
            else:
                attempts.append(f"not-writable:{primary_dir}")
        except Exception as e:
            attempts.append(f"error:{primary_dir}:{e}")

        if not written:
            for fb in fallback_dirs:
                try:
                    fb = fb.resolve()
                    if is_writable_dir(fb):
                        csv_raw_fb = fb / f"reconstructed_table_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        final_xlsx_fb = fb / f"final_structured_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        safe_write_excel(preliminary_df, csv_raw_fb, sheet_name='Reconstructed')
                        safe_write_excel(final_structured_df, final_xlsx_fb, sheet_name='Final')
                        print(f"[!] Écriture+fallback réussis dans : {fb}")
                        csv_raw = csv_raw_fb
                        final_xlsx_path = final_xlsx_fb
                        written = True
                        attempts.append(str(fb))
                        break
                    else:
                        attempts.append(f"not-writable:{fb}")
                except Exception as ex:
                    attempts.append(f"error:{fb}:{ex}")

        if not written:
            print("Erreur: Impossible d'écrire dans le dossier demandé et dans les dossiers de secours.")
            print("Emplacements testés:", attempts)
            raise PermissionError(f"Impossible d'écrire les fichiers Excel (checked: {attempts})")
    else:
        # fallback: write CSVs (maintain previous behaviour)
        raw_csv_path = Path(out_dir) / "reconstructed_table.csv"
        final_csv_path = Path(out_dir) / "final_structured.csv"
        raw_opts = {"index": False, "encoding": "utf-8-sig"}
        final_opts = {"index": False, "sep": ";", "encoding": "utf-8-sig", "quoting": csv.QUOTE_ALL}
        attempts = []
        written = False
        try:
            if is_writable_dir(primary_dir):
                safe_write_csv(preliminary_df, raw_csv_path, **raw_opts)
                safe_write_csv(final_structured_df, final_csv_path, **final_opts)
                written = True
                attempts.append(str(primary_dir))
            else:
                attempts.append(f"not-writable:{primary_dir}")
        except Exception as e:
            attempts.append(f"error:{primary_dir}:{e}")

        if not written:
            for fb in fallback_dirs:
                try:
                    fb = fb.resolve()
                    if is_writable_dir(fb):
                        raw_csv_fb = fb / f"reconstructed_table_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                        final_csv_fb = fb / f"final_structured_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                        safe_write_csv(preliminary_df, raw_csv_fb, **raw_opts)
                        safe_write_csv(final_structured_df, final_csv_fb, **final_opts)
                        print(f"[!] Écriture+fallback réussis dans : {fb}")
                        raw_csv_path = raw_csv_fb
                        final_csv_path = final_csv_fb
                        written = True
                        attempts.append(str(fb))
                        break
                    else:
                        attempts.append(f"not-writable:{fb}")
                except Exception as ex:
                    attempts.append(f"error:{fb}:{ex}")

        if not written:
            print("Erreur: Impossible d'écrire dans le dossier demandé et dans les dossiers de secours.")
            print("Emplacements testés:", attempts)
            raise PermissionError(f"Impossible d'écrire les CSV (checked: {attempts})")

    # debug images
    try:
        cv2.imwrite(os.path.join(out_dir, "debug_grid_overlay.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
        cv2.imwrite(os.path.join(out_dir, "debug_hv_lines.png"), cv2.cvtColor(lines_vis, cv2.COLOR_RGB2BGR))
    except Exception:
        pass

    print(f"[+] CSV raw (prelim): {csv_raw}")
    print(f"[+] Final structured file: {final_xlsx_path}")
    print(f"[+] Debug images in {out_dir}")

    # Optionally visualize (if requested)
    if visualize:
        import matplotlib.pyplot as plt
        plt.figure(figsize=(14,10))
        plt.subplot(2,2,1); plt.imshow(orig); plt.title("Original"); plt.axis('off')
        plt.subplot(2,2,2); plt.imshow(bin_img, cmap='gray'); plt.title("Binarized (inverted)"); plt.axis('off')
        plt.subplot(2,2,3); plt.imshow(lines_vis); plt.title("Detected horizontal (green) & vertical (blue)"); plt.axis('off')
        plt.subplot(2,2,4); plt.imshow(overlay); plt.title("Cell candidates (red boxes)"); plt.axis('off')
        plt.tight_layout(); plt.show()
        print("Aperçu DataFrame reconstruit (premières lignes) ---")
        print(preliminary_df.head(30).to_string(index=False))

    return {
        "csv_raw": str(csv_raw),
        "final_csv": str(final_xlsx_path),
        "debug_images": {
            "overlay": os.path.join(out_dir, "debug_overlay_cells.png"),
            "lines": os.path.join(out_dir, "debug_lines_hv.png"),
            "binarized": os.path.join(out_dir, "debug_binarized.png"),
            "grid_overlay": os.path.join(out_dir, "debug_grid_overlay.png"),
            "cells_dir": os.path.join(out_dir, "cells")
        },
        "final_df": final_structured_df
    }

    # debug images
    try:
        cv2.imwrite(os.path.join(out_dir, "debug_grid_overlay.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
        cv2.imwrite(os.path.join(out_dir, "debug_hv_lines.png"), cv2.cvtColor(lines_vis, cv2.COLOR_RGB2BGR))
    except Exception:
        pass

    print(f"[+] CSV raw (prelim): {csv_raw}")
    print(f"[+] Final structured CSV: {final_csv_path}")
    print(f"[+] Debug images in {out_dir}")

    # Optionally visualize (if requested)
    if visualize:
        import matplotlib.pyplot as plt
        plt.figure(figsize=(14,10))
        plt.subplot(2,2,1); plt.imshow(orig); plt.title("Original"); plt.axis('off')
        plt.subplot(2,2,2); plt.imshow(bin_img, cmap='gray'); plt.title("Binarized (inverted)"); plt.axis('off')
        plt.subplot(2,2,3); plt.imshow(lines_vis); plt.title("Detected horizontal (green) & vertical (blue)"); plt.axis('off')
        plt.subplot(2,2,4); plt.imshow(overlay); plt.title("Cell candidates (red boxes)"); plt.axis('off')
        plt.tight_layout(); plt.show()
        print("\n--- Aperçu DataFrame reconstruit (premières lignes) ---")
        print(preliminary_df.head(30).to_string(index=False))

    return {
        "csv_raw": str(csv_raw),
        "final_csv": str(final_csv_path),
        "debug_images": {
            "overlay": os.path.join(out_dir, "debug_overlay_cells.png"),
            "lines": os.path.join(out_dir, "debug_lines_hv.png"),
            "binarized": os.path.join(out_dir, "debug_binarized.png"),
            "grid_overlay": os.path.join(out_dir, "debug_grid_overlay.png"),
            "cells_dir": os.path.join(out_dir, "cells")
        },
        "final_df": final_structured_df
    }

# ---------- CLI ----------

def parse_args():
    p = argparse.ArgumentParser(description="OCR pipeline table -> CSV (Tesseract) - improved")
    p.add_argument("--pdf", required=True, help="Chemin vers le fichier PDF ou image")
    p.add_argument("--page", type=int, default=0, help="Numéro de la page (0-index)")
    p.add_argument("--zoom", type=float, default=3.0, help="Facteur de rendu (3.0 recommandé)")
    p.add_argument("--out_dir", default="./output", help="Dossier de sortie")
    p.add_argument("--visualize", action="store_true", help="Afficher matplotlib visualisations (nécessite Qt/Tk)")
    p.add_argument("--no_save_cells", action="store_true", help="Ne pas sauvegarder les images de cellules (gain d'espace)")
    p.add_argument("--psm", type=int, default=6, help="Tesseract PSM (6 recommended, try 4,11 if necessary)")
    return p.parse_args()


if __name__ == "__main__":
    import sys
    import matplotlib
    try:
        args = parse_args()
    except SystemExit:
        print(" Erreur arguments manquants ou incorrects. ")
        sys.exit(2)

    if args.visualize:
        try:
            matplotlib.use('Qt5Agg')
        except Exception:
            try:
                matplotlib.use('TkAgg')
            except Exception:
                matplotlib.use('Agg')
    else:
        matplotlib.use('Agg')

    ensure_dir(args.out_dir)
    out = pipeline(args.pdf, page=args.page, zoom=args.zoom, out_dir=args.out_dir,
                   visualize=args.visualize, debug_save_cells=(not args.no_save_cells), psm=args.psm)

    print(" Résumé ")
    print("Final structured CSV :", out.get("final_csv"))
    print("Raw reconstructed CSV :", out.get("csv_raw"))
    print("Debug images:")
    for k,v in out.get("debug_images", {}).items():
        print(f" - {k}: {v}")
    print("Terminé.")
