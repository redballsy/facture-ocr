import sys
import os
import re
import io
import json
import zipfile
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, render_template, request, jsonify, send_file

try:
    import pytesseract
    from PIL import Image
    import cv2
    import numpy as np
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from pdf2image import convert_from_bytes
except ImportError as e:
    print(f"❌ Dépendance manquante: {e}")
    print("Lance: pip install flask pytesseract pillow opencv-python openpyxl pdf2image")
    sys.exit(1)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max pour 100 fichiers
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "gesfi-secret-key-2026")

EXTENSIONS = {"jpg", "jpeg", "png", "bmp", "tiff", "tif", "webp", "gif", "pdf"}

# ========== COULEURS GESFI GROUP ==========
COULEURS_GESFI = {
    "primary": "#1B3A5C",
    "secondary": "#2E5A88",
    "accent": "#D4AF37",
    "accent_light": "#F5E6B8",
    "text_dark": "#1A1A1A",
    "text_light": "#FFFFFF",
    "background": "#F8F9FA",
    "success": "#28A745",
    "warning": "#FFC107",
    "error": "#DC3545",
    "border": "#DEE2E6",
}

COULEURS_EXCEL = {
    "header_bg": "1B3A5C",
    "header_font": "FFFFFF",
    "ligne_paire": "F0F4F8",
    "total_bg": "2E5A88",
    "total_font": "FFFFFF",
}

CHEMINS_TESSERACT = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    "/usr/bin/tesseract",
    "/usr/local/bin/tesseract",
]

def configurer_tesseract():
    for chemin in CHEMINS_TESSERACT:
        if Path(chemin).exists():
            pytesseract.pytesseract.tesseract_cmd = chemin
            return True
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False

TESSERACT_OK = configurer_tesseract()

# ── Prétraitement image ───────────────────────────────────────────────────────

def pretraiter_image(img_bytes: bytes) -> Image.Image:
    arr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
    h, w = img.shape[:2]
    if w < 1200:
        scale = 1200 / w
        img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.fastNlMeansDenoising(gray, h=10)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    thresh = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10
    )
    return Image.fromarray(thresh)

def extraire_texte_image(img_bytes: bytes) -> str:
    img = pretraiter_image(img_bytes)
    config = "--oem 3 --psm 6"
    try:
        return pytesseract.image_to_string(img, lang="fra+eng", config=config)
    except Exception:
        try:
            return pytesseract.image_to_string(img, lang="eng", config=config)
        except Exception:
            return ""

def traiter_pdf(img_bytes: bytes) -> str:
    """Convertit un PDF en images et extrait le texte"""
    try:
        images = convert_from_bytes(img_bytes, dpi=300)
        texte_complet = ""
        
        for i, image in enumerate(images):
            img_byte_arr = io.BytesIO()
            image.save(img_byte_arr, format='PNG')
            img_byte_arr.seek(0)
            texte = extraire_texte_image(img_byte_arr.getvalue())
            texte_complet += f"\n--- Page {i+1} ---\n{texte}\n"
        
        return texte_complet
    except Exception as e:
        print(f"Erreur lors du traitement PDF: {e}")
        return ""

def extraire_texte(img_bytes: bytes, filename: str) -> str:
    """Extrait le texte d'une image ou d'un PDF selon l'extension"""
    extension = filename.lower().split('.')[-1] if '.' in filename else ''
    
    if extension == 'pdf':
        return traiter_pdf(img_bytes)
    else:
        return extraire_texte_image(img_bytes)

# ── Fonctions utilitaires ─────────────────────────────────────────────────────

def net(s):
    return s.strip() if s else ""

def chercher(pattern, texte, groupe=1, defaut=""):
    m = re.search(pattern, texte, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return net(m.group(groupe)) if m else defaut

def nettoyer_montant(montant_str):
    if not montant_str:
        return ""
    montant = re.sub(r'[^\d,\.\s]', '', montant_str)
    montant = re.sub(r'\s+', '', montant)
    montant = montant.replace(',', '.')
    return montant

def detecter_type_facture(texte: str) -> str:
    """Détecte si c'est une facture d'achat ou standard"""
    texte_lower = texte.lower()
    if "facture d'achat" in texte_lower or "po-" in texte_lower or "livraison prévue" in texte_lower or "fournisseur" in texte_lower:
        return "achat"
    return "standard"

# ── Extraction FACTURE STANDARD (facture.png) ─────────────────────────────────

def extraire_facture_standard(texte: str) -> list:
    """Extraction pour le template facture standard (vente)"""
    donnees = []
    
    num_facture = chercher(r'FACTURE\s+([A-Z0-9\-]+)', texte)
    if not num_facture:
        num_facture = chercher(r'(?:INV|N°|Numéro)[\s\-:]*([A-Z0-9\-]+)', texte)
    if num_facture:
        donnees.append(("N° FACTURE", num_facture))
    
    date_emission = chercher(r'Date d\'émission:\s*([^\n]+)', texte)
    if date_emission:
        donnees.append(("Date d'émission", date_emission))
    
    date_echeance = chercher(r'Date d\'échéance:\s*([^\n]+)', texte)
    if date_echeance:
        donnees.append(("Date d'échéance", date_echeance))
    
    source = chercher(r'Source\s*([^\n]+)', texte)
    if not source:
        source = chercher(r'(\d+\s+BP\s+[A-Z]+,\s*[A-Z]+,\s*Côte', texte)
    if source:
        donnees.append(("Source", source))
    
    rccm = chercher(r'(RCCM[\s\-]+[A-Z0-9\-]+)', texte)
    if rccm:
        donnees.append(("RCCM", rccm))
    
    capital1 = chercher(r'CAPITAL\s+(\d[\d\s]+)', texte)
    if capital1:
        donnees.append(("CAPITAL", capital1))
    
    client = chercher(r'CLIENT\s+([^\n]+)', texte)
    if not client:
        client = chercher(r'Client[:\s]+([^\n]+)', texte)
    if client:
        donnees.append(("Client", client))
    
    desc = chercher(r'Tablette|Description\s+([^\n]+)', texte)
    if not desc:
        desc = "Tablette"
    if desc:
        donnees.append(("Description", desc))
    
    qte = chercher(r'Tablette\s+(\d+)', texte)
    if not qte:
        qte = chercher(r'Qté\s+(\d+)', texte)
    if qte:
        donnees.append(("Quantité", qte))
    
    prix = chercher(r'Tablette\s+\d+\s+([\d\s\.,]+)', texte)
    if not prix:
        prix = chercher(r'Prix HT\s+([\d\s\.,]+)', texte)
    if prix:
        donnees.append(("Prix HT", nettoyer_montant(prix) + " FCFA"))
    
    tva_pct = chercher(r'TVA\s+(\d+%?)', texte)
    if tva_pct:
        donnees.append(("TVA %", tva_pct))
    
    total_ligne = chercher(r'Tablette\s+\d+\s+[\d\s\.,]+\s+\d+%?\s+([\d\s\.,]+)', texte)
    if total_ligne:
        donnees.append(("Total ligne", nettoyer_montant(total_ligne) + " FCFA"))
    
    sous_total = chercher(r'Sous-total HT:\s*([\d\s\.,]+)', texte)
    if sous_total:
        donnees.append(("Sous-total HT", nettoyer_montant(sous_total) + " FCFA"))
    
    tva_montant = chercher(r'TVA:\s*([\d\s\.,]+)', texte)
    if tva_montant:
        donnees.append(("TVA (montant)", nettoyer_montant(tva_montant) + " FCFA"))
    
    total_ttc = chercher(r'Total TTC:\s*([\d\s\.,]+)', texte)
    if total_ttc:
        donnees.append(("Total TTC", nettoyer_montant(total_ttc) + " FCFA"))
    
    capital2 = chercher(r'CAPITAL\s+(\d[\d\s]+)(?=.*$)', texte)
    if capital2 and capital2 != capital1:
        donnees.append(("CAPITAL (second)", capital2))
    
    paiement = chercher(r'PAIEMENT\s*:\s*([^\n]+)', texte)
    if paiement:
        donnees.append(("Mode de paiement", paiement))
    
    return donnees

# ── Extraction FACTURE ACHAT (facture2.png) ─────────────────────────────────

def extraire_facture_achat(texte: str) -> list:
    """Extraction pour le template facture d'achat"""
    donnees = []
    
    num_facture = chercher(r'N°\s*([A-Z0-9\-]+)', texte)
    if not num_facture:
        num_facture = chercher(r'PO-([A-Z0-9\-]+)', texte)
    if num_facture:
        donnees.append(("N° FACTURE", num_facture))
    
    date = chercher(r'Date:\s*([^\n]+)', texte)
    if date:
        donnees.append(("Date", date))
    
    livraison = chercher(r'Livraison prévue:\s*([^\n]+)', texte)
    if livraison:
        donnees.append(("Livraison prévue", livraison))
    
    source = chercher(r'Source\s*([^\n]+)', texte)
    if source:
        donnees.append(("Source", source))
    
    fournisseur_nom = chercher(r'FOURNISSEUR\s*([^\n]+)', texte)
    if fournisseur_nom:
        donnees.append(("Fournisseur Nom", fournisseur_nom))
    
    fournisseur_adresse = chercher(r'FOURNISSEUR\s*[^\n]+\s*([^\n]+)', texte, 1, "")
    if fournisseur_adresse:
        donnees.append(("Fournisseur Adresse", fournisseur_adresse))
    
    fournisseur_tel = chercher(r'Tél:\s*([^\n]+)', texte)
    if fournisseur_tel:
        donnees.append(("Fournisseur Téléphone", fournisseur_tel))
    
    fournisseur_email = chercher(r'Email:\s*([^\n]+)', texte)
    if fournisseur_email:
        donnees.append(("Fournisseur Email", fournisseur_email))
    
    desc = chercher(r'Imprimante de bureau|Description\s+([^\n]+)', texte)
    if not desc:
        desc = chercher(r'\| (.+?) \|', texte)
    if desc:
        donnees.append(("Description", desc))
    
    qte = chercher(r'Imprimante\s+([\d\.,]+)', texte)
    if not qte:
        qte = chercher(r'Qté\s+([\d\.,]+)', texte)
    if qte:
        donnees.append(("Quantité", qte))
    
    prix = chercher(r'Imprimante\s+[\d\.,]+\s+([\d\s\.,]+)', texte)
    if prix:
        donnees.append(("Prix Unitaire", nettoyer_montant(prix) + " FCFA"))
    
    tva_pct = chercher(r'TVA\s+(\d+%?)', texte)
    if tva_pct:
        donnees.append(("TVA %", tva_pct))
    
    total_ligne = chercher(r'Imprimante\s+[\d\.,]+\s+[\d\s\.,]+\s+\d+%?\s+([\d\s\.,]+)', texte)
    if total_ligne:
        donnees.append(("Total HT ligne", nettoyer_montant(total_ligne) + " FCFA"))
    
    sous_total = chercher(r'Sous-total HT:\s*([\d\s\.,]+)', texte)
    if sous_total:
        donnees.append(("Sous-total HT", nettoyer_montant(sous_total) + " FCFA"))
    
    tva_montant = chercher(r'TVA:\s*([\d\s\.,]+)', texte)
    if tva_montant:
        donnees.append(("TVA (montant)", nettoyer_montant(tva_montant) + " FCFA"))
    
    total_ttc = chercher(r'Total TTC:\s*([\d\s\.,]+)', texte)
    if total_ttc:
        donnees.append(("Total TTC", nettoyer_montant(total_ttc) + " FCFA"))
    
    return donnees

# ── Parser principal ─────────────────────────────────────────────────────────

def parser_facture_complet(texte: str, nom_fichier: str) -> dict:
    """Parse toute facture selon son type détecté"""
    type_facture = detecter_type_facture(texte)
    
    if type_facture == "achat":
        donnees_liste = extraire_facture_achat(texte)
        type_label = "ACHAT"
    else:
        donnees_liste = extraire_facture_standard(texte)
        type_label = "STANDARD"
    
    return {
        "nom_fichier": nom_fichier,
        "type": type_label,
        "donnees": donnees_liste,
        "texte_brut": texte
    }

# ── Génération Excel (2 colonnes seulement) ──────────────────────────────────

def generer_excel_structure(donnees: dict, nom_fichier: str) -> bytes:
    """Génère Excel avec 2 colonnes: LIBELLÉ | VALEUR"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Facture"
    
    ws.cell(row=1, column=1, value="LIBELLÉ")
    ws.cell(row=1, column=2, value="VALEUR")
    
    for col in [1, 2]:
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, color=COULEURS_EXCEL["header_font"], size=11)
        cell.fill = PatternFill("solid", start_color=COULEURS_EXCEL["header_bg"])
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 45
    
    row = 2
    
    ws.cell(row=row, column=1, value="TYPE FACTURE")
    ws.cell(row=row, column=2, value=donnees.get("type", "STANDARD"))
    row += 1
    
    for libelle, valeur in donnees.get("donnees", []):
        if valeur:
            ws.cell(row=row, column=1, value=libelle)
            ws.cell(row=row, column=2, value=str(valeur))
            row += 1
    
    for r in range(2, row):
        for c in [1, 2]:
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c == 1:
                cell.font = Font(bold=True)
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Traitement batch ─────────────────────────────────────────────────────────

def traiter_un_fichier(fichier_info: tuple) -> dict:
    idx, fichier = fichier_info
    try:
        img_bytes = fichier.read()
        fichier.seek(0)
        
        texte = extraire_texte(img_bytes, fichier.filename)
        
        if not texte.strip():
            return {"index": idx, "nom": fichier.filename, "erreur": "Aucun texte détecté"}
        
        donnees = parser_facture_complet(texte, fichier.filename)
        return {"index": idx, "nom": fichier.filename, "succes": True, "donnees": donnees}
    except Exception as e:
        return {"index": idx, "nom": fichier.filename, "erreur": str(e)}

# ── Routes Flask ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html", tesseract_ok=TESSERACT_OK, couleurs=COULEURS_GESFI)

@app.route("/analyser_batch", methods=["POST"])
def analyser_batch():
    if not TESSERACT_OK:
        return jsonify({"erreur": "Tesseract non installé"}), 500
    
    if "fichiers" not in request.files:
        return jsonify({"erreur": "Aucun fichier reçu"}), 400
    
    fichiers = request.files.getlist("fichiers")
    
    if len(fichiers) > 100:
        return jsonify({"erreur": "Maximum 100 fichiers autorisés"}), 400
    
    if len(fichiers) == 0:
        return jsonify({"erreur": "Aucun fichier sélectionné"}), 400
    
    resultats = []
    fichiers_a_traiter = [(i, f) for i, f in enumerate(fichiers) if f.filename]
    
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(traiter_un_fichier, ft): ft for ft in fichiers_a_traiter}
        for future in as_completed(futures):
            resultats.append(future.result())
    
    resultats.sort(key=lambda x: x.get("index", 0))
    
    return jsonify({
        "total": len(resultats),
        "succes": len([r for r in resultats if r.get("succes")]),
        "erreurs": len([r for r in resultats if r.get("erreur")]),
        "resultats": resultats
    })

@app.route("/telecharger_batch", methods=["POST"])
def telecharger_batch():
    data = request.get_json()
    if not data or "resultats" not in data:
        return jsonify({"erreur": "Données manquantes"}), 400
    
    resultats = data.get("resultats", [])
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for resultat in resultats:
            if resultat.get("succes") and resultat.get("donnees"):
                donnees = resultat["donnees"]
                nom_original = Path(resultat["nom"]).stem
                excel_bytes = generer_excel_structure(donnees, nom_original)
                zip_file.writestr(f"{nom_original}.xlsx", excel_bytes)
    
    zip_buffer.seek(0)
    
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"factures_extraites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    )

@app.route("/telecharger_unique", methods=["POST"])
def telecharger_unique():
    donnees = request.get_json()
    if not donnees:
        return jsonify({"erreur": "Données manquantes"}), 400
    
    nom_fichier = donnees.get("nom_fichier", "facture")
    excel_bytes = generer_excel_structure(donnees, nom_fichier)
    nom_sortie = f"{Path(nom_fichier).stem}.xlsx"
    
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=nom_sortie,
    )

@app.route("/sante")
def sante():
    return jsonify({
        "status": "ok",
        "tesseract": TESSERACT_OK,
        "version": "3.0.0",
        "marque": "GESFI GROUP",
        "batch_max": 100,
        "formats_supportes": ["JPG", "PNG", "BMP", "TIFF", "WEBP", "PDF"],
        "timestamp": datetime.now().isoformat()
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)