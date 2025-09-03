import os, fitz, base64, json, re
import google.generativeai as genai
from dotenv import load_dotenv
import pandas as pd


# ----------------- Config multi-API -----------------
load_dotenv()
API_KEYS = [
    os.getenv("GEMINI_API_KEY1"),
    os.getenv("GEMINI_API_KEY2"),
    os.getenv("GEMINI_API_KEY3")
]

# Retire les clés None
API_KEYS = [k for k in API_KEYS if k]
if not API_KEYS:
    raise RuntimeError("⚠️ Mets au moins une clé dans .env : GEMINI_API_KEY1, GEMINI_API_KEY2")

current_key_index = 0

def init_genai():
    """Initialise Gemini avec la clé courante"""
    global current_key_index
    genai.configure(api_key=API_KEYS[current_key_index])

def switch_key():
    """Bascule vers la clé suivante en cas de limite atteinte"""
    global current_key_index
    current_key_index = (current_key_index + 1) % len(API_KEYS)
    print(f"⚠️ Bascule vers la clé {current_key_index+1}")
    init_genai()

# Initialisation
init_genai()

#---------------PARTIE PROMPT-----------------------

PROMPT_BILAN_TABLEAU = """
Tu es un assistant OCR intelligent. Tu reçois en entrée soit :
- une ou plusieurs images (issues d’un scan PDF ou d’une photo) contenant un tableau scanné de bilan comptable en français,
- soit un PDF scanné (dans ce cas, commence par convertir chaque page du PDF en image avant de traiter).

La première partie de l'image peut contenir du texte descriptif ou des informations inutiles. **Ignore tout le texte avant le tableau**.
La qualité peut être faible et contenir du bruit (caractères illisibles, symboles parasites).
Ta tâche est d’extraire uniquement le **tableau du bilan Actif ou Passif** et de produire du JSON strictement valide.

### Étapes de traitement :
1. Si tu reçois un PDF scanné, convertis chaque page en image avant de commencer l’OCR.
2. Localise le tableau et ignore tout texte en dehors du tableau.
3. Nettoie le texte OCR :
   - Supprime tous les caractères parasites (ex : â€” ‚ ¨ | [ ] { } / \ _ ° * # etc.).
   - Supprime les fragments incompréhensibles.
   - Corrige les accents si possible (é, è, à, ç).
   - Réduis plusieurs espaces en un seul.
4. Normalise les nombres :
   - Garde uniquement les chiffres et la virgule/point.
   - Convertis tous les nombres au format `123456.78` (séparateur décimal = point).
   - Si une case est vide ou illisible → valeur = "".

### Structure JSON attendue :

**Pour le Bilan Actif :**
[
  {
    "Type_Tableau": "Bilan_Actif",
    "Sous_Categorie": "IMMOBILISATIONS INCORPORELLES",
    "Rubrique": "Frais d'établissement",
    "Montant_Brut": "209066.09",
    "Amortissements_Provisions": "209066.09",
    "Net_Exercice": "0.00",
    "Net_Exercice_Prec": "100.50",
    "Commentaires": ""
  }
]

**Pour le Bilan Passif :**
[
  {
    "Type_Tableau": "Bilan_Passif",
    "Sous_Categorie": "CAPITAUX PROPRES",
    "Rubrique": "Capital social ou personnel",
    "Montant_Brut": "",
    "Amortissements_Provisions": "",
    "Net_Exercice": "",
    "Net_Exercice_Prec": "",
    "Commentaires": ""
  }
]

### Contraintes :
- Toujours renvoyer uniquement une **liste JSON valide** (pas de texte autour).
- Si un champ est absent → mets "".
- Ne traite que les lignes présentes dans le tableau et ignore tout le reste.
- Lors de l’export Excel, **sépare chaque type de bilan dans un onglet distinct** (par exemple : feuille "Actif" et feuille "Passif").
"""


def safe_json_extract(text):
    """Nettoie la réponse Gemini pour récupérer un JSON valide (liste)."""
    if not text:
        return []
    # debug: affiche un extrait quand on échoue pour déboguer
    match = re.search(r"\[.*\]", text, re.S)
    if not match:
        print("⚠️ Aucun JSON détecté dans la réponse Gemini (preview):", text[:500])
        return []
    raw = match.group(0)
    try:
        data = json.loads(raw)
        return data if isinstance(data, list) else []
    except Exception as e:
        print("⚠️ JSON non valide après nettoyage:", e)
        print("Réponse brute Gemini (500 chars):", text[:500])
        return []


#AJOUTER CALL GEMINI SI LE API KEY ATTEIND SA LIMITE------
def call_gemini(content, model_name="gemini-1.5-flash"):
    """Appelle Gemini avec fallback automatique"""
    global current_key_index
    init_genai()
    try:
        model = genai.GenerativeModel(model_name)
        resp = model.generate_content(content)
        text = getattr(resp, "text", None) or getattr(resp, "output_text", None) or str(resp)
        return text
    except Exception as e:
        if "429" in str(e) or "quota" in str(e).lower():
            print("🚨 Limite atteinte → changement de clé")
            switch_key()
            return call_gemini(content, model_name)
        else:
            raise e



def extract_images_from_pdf(pdf_path, zoom=2):
    """Convertit chaque page PDF en image (png bytes). zoom par défaut 2 => taille raisonnable."""
    import fitz
    doc = fitz.open(pdf_path)
    img_bytes_list = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        img_bytes_list.append(pix.tobytes("png"))
    doc.close()
    return img_bytes_list


def ocr_pages_with_gemini(img_bytes_list, model_name="gemini-1.5-flash"):
    """
    Envoie un batch à Gemini. Si le batch renvoie vide ou JSON invalide,
    retente image par image (fallback).
    """
    # Prépare le payload attendu par call_gemini : [PROMPT, {mime_type, data}, ...]
    batch_payload = [PROMPT_BILAN_TABLEAU]
    for b in img_bytes_list:
        batch_payload.append({"mime_type": "image/png", "data": base64.b64encode(b).decode("utf-8")})

    text = call_gemini(batch_payload, model_name=model_name)
    rows = safe_json_extract(text)
    if rows:
        return rows

    # Fallback : envoi image par image (parfois plus fiable)
    fallback_rows = []
    for i, b in enumerate(img_bytes_list):
        print(f"⚠️ Fallback page {i+1}/{len(img_bytes_list)} envoi séparé...")
        try:
            payload = [PROMPT_BILAN_TABLEAU, {"mime_type": "image/png", "data": base64.b64encode(b).decode("utf-8")}]
            t = call_gemini(payload, model_name=model_name)
            r = safe_json_extract(t)
            if r:
                fallback_rows.extend(r)
        except Exception as e:
            print("Erreur sur page fallback:", e)
    return fallback_rows


def process_inputs(inputs, batch_size=5, out_xlsx=None):
    """
    Traite un PDF scanné ou une liste d'images
    inputs : chemin vers un PDF, un dossier d'images, ou une liste de fichiers image
    """
    all_imgs = []

    # Cas 1 : un PDF
    if isinstance(inputs, str) and inputs.lower().endswith(".pdf"):
        all_imgs = extract_images_from_pdf(inputs)

    # Cas 2 : un dossier d'images
    elif isinstance(inputs, str) and os.path.isdir(inputs):
        for fname in sorted(os.listdir(inputs)):
            if fname.lower().endswith((".png", ".jpg", ".jpeg")):
                with open(os.path.join(inputs, fname), "rb") as f:
                    all_imgs.append(f.read())

    # Cas 3 : une liste de fichiers image
    elif isinstance(inputs, list):
        for img_path in inputs:
            with open(img_path, "rb") as f:
                all_imgs.append(f.read())

    else:
        raise ValueError("Entrée non reconnue. Donne un PDF, un dossier d'images ou une liste de fichiers image.")

    if not all_imgs:
        raise RuntimeError("⚠️ Aucune image trouvée pour traitement.")

    # Nom de sortie
    if out_xlsx is None:
        base = os.path.splitext(os.path.basename(inputs if isinstance(inputs, str) else "bilan"))[0]
        out_xlsx = f"{base}_bilan_gemini.xlsx"

    all_rows = []
    for batch_start in range(0, len(all_imgs), batch_size):
        batch_end = min(batch_start + batch_size, len(all_imgs))
        print(f"📄 Traitement images {batch_start+1} à {batch_end} / {len(all_imgs)} ...")
        batch_imgs = all_imgs[batch_start:batch_end]
        rows = ocr_pages_with_gemini(batch_imgs)
        if rows:
            all_rows.extend(rows)

    df = pd.DataFrame(all_rows)
    if not df.empty:
        for col in ["Sous_Categorie", "Rubrique"]:
            if col in df:
                df[col] = df[col].astype(str).str.strip()
        df = df.dropna(how="all")
        df = df[~((df["Rubrique"] == "") & (df["Montant_Brut"] == ""))]

        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            df_actif = df[df["Type_Tableau"] == "Bilan_Actif"]
            df_passif = df[df["Type_Tableau"] == "Bilan_Passif"]
            if not df_actif.empty:
                df_actif.to_excel(writer, sheet_name="Actif", index=False)
            if not df_passif.empty:
                df_passif.to_excel(writer, sheet_name="Passif", index=False)

        print(f"✅ Extraction Bilan → {out_xlsx} ({len(df)} lignes)")
        print("Excel généré :", out_xlsx)
        print(f"DEBUG: {len(all_imgs)} images extraites. Taille première image: {len(all_imgs[0]) if all_imgs else 0} bytes")
        return out_xlsx
    
    else:
        print("⚠️ Aucun résultat exploitable extrait")
        return None
