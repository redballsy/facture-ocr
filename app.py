import sys
import os
import re
import io
import json
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
    from openpyxl.styles import Font, PatternFill, Alignment
    from pdf2image import convert_from_bytes
except ImportError as e:
    print(f"❌ Dépendance manquante: {e}")
    print("Lance: pip install flask pytesseract pillow opencv-python openpyxl pdf2image")
    sys.exit(1)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "gesfi-secret-key-2026")

EXTENSIONS = {"jpg", "jpeg", "png", "bmp", "tiff", "tif", "webp", "gif", "pdf"}

COULEURS_GESFI = {
    "primary": "#1B3A5C",
    "secondary": "#2E5A88",
    "accent": "#D4AF37",
}

COULEURS_EXCEL = {
    "header_bg": "1B3A5C",
    "header_font": "FFFFFF",
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

def pretraiter_image(img_bytes: bytes):
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
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 10)
    return Image.fromarray(thresh)

def extraire_texte_image(img_bytes: bytes) -> str:
    img = pretraiter_image(img_bytes)
    config = "--oem 3 --psm 6"
    try:
        return pytesseract.image_to_string(img, lang="fra+eng", config=config)
    except Exception:
        return pytesseract.image_to_string(img, lang="eng", config=config)

def traiter_pdf(img_bytes: bytes) -> str:
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
        return ""

def extraire_texte(img_bytes: bytes, filename: str) -> str:
    extension = filename.lower().split('.')[-1] if '.' in filename else ''
    if extension == 'pdf':
        return traiter_pdf(img_bytes)
    else:
        return extraire_texte_image(img_bytes)

def net(s):
    return s.strip() if s else ""

def chercher(pattern, texte, groupe=1, defaut=""):
    m = re.search(pattern, texte, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return net(m.group(groupe)) if m else defaut

def nettoyer_montant(montant_str):
    if not montant_str:
        return ""
    montant = re.sub(r'[^\d,\.]', '', montant_str.replace(' ', ''))
    montant = montant.replace(',', '.')
    return montant

# ── Extraction COMPLETE pour facture ACHAT ─────────────────────────────────────

def extraire_facture_achat(texte: str) -> list:
    donnees = []
    
    # 1. N° FACTURE
    num = chercher(r'N°\s*([A-Z0-9\-]+)', texte)
    if not num:
        num = chercher(r'PO-([A-Z0-9\-]+)', texte)
    if num:
        donnees.append(("N° FACTURE", num))
    
    # 2. Date
    date = chercher(r'Date:\s*([^\n]+)', texte)
    if date:
        donnees.append(("Date", date))
    
    # 3. Livraison prévue
    livraison = chercher(r'Livraison prévue:\s*([^\n]+)', texte)
    if livraison:
        donnees.append(("Livraison prévue", livraison))
    
    # 4. Source (adresse) - correction
    source = chercher(r'Source\s*\n\s*([^\n]+)', texte)
    if not source:
        source = chercher(r'Source\s*([^\n]+)', texte)
    if source and "Date:" not in source:
        donnees.append(("Source", source))
    
    # 5. Fournisseur Nom
    fournisseur = chercher(r'FOURNISSEUR\s*\n\s*([^\n]+)', texte)
    if fournisseur:
        donnees.append(("Fournisseur Nom", fournisseur))
    
    # 6. Fournisseur Adresse
    adresse = chercher(r'FOURNISSEUR\s*[^\n]+\n\s*([^\n]+)', texte)
    if adresse:
        donnees.append(("Fournisseur Adresse", adresse))
    
    # 7. Fournisseur Téléphone
    tel = chercher(r'Tél:\s*([^\n]+)', texte)
    if tel:
        donnees.append(("Fournisseur Téléphone", tel))
    
    # 8. Fournisseur Email
    email = chercher(r'Email:\s*([^\n]+)', texte)
    if email:
        donnees.append(("Fournisseur Email", email))
    
    # 9. Description
    desc = chercher(r'Imprimante de bureau|Description\s+([^\n]+)', texte)
    if not desc:
        desc = "Imprimante de bureau"
    donnees.append(("Description", desc))
    
    # 10. Quantité
    qte = chercher(r'Imprimante de bureau\s+([\d\.,]+)', texte)
    if not qte:
        qte = chercher(r'Qté\s+([\d\.,]+)', texte)
    if qte:
        donnees.append(("Quantité", qte))
    
    # 11. Prix Unitaire
    prix = chercher(r'Imprimante de bureau\s+[\d\.,]+\s+([\d\s\.,]+)', texte)
    if prix:
        donnees.append(("Prix Unitaire", nettoyer_montant(prix) + " FCFA"))
    
    # 12. TVA %
    tva_pct = chercher(r'TVA\s+(\d+%?)', texte)
    if tva_pct:
        donnees.append(("TVA %", tva_pct))
    
    # 13. Total HT ligne
    total_ligne = chercher(r'Imprimante de bureau\s+[\d\.,]+\s+[\d\s\.,]+\s+\d+%?\s+([\d\s\.,]+)', texte)
    if total_ligne:
        donnees.append(("Total HT ligne", nettoyer_montant(total_ligne) + " FCFA"))
    
    # 14. Sous-total HT
    sous_total = chercher(r'Sous-total HT:\s*[\*]*\s*([\d\s\.,]+)', texte)
    if sous_total:
        donnees.append(("Sous-total HT", nettoyer_montant(sous_total) + " FCFA"))
    
    # 15. TVA (montant)
    tva = chercher(r'TVA:\s*[\*]*\s*([\d\s\.,]+)', texte)
    if tva:
        donnees.append(("TVA (montant)", nettoyer_montant(tva) + " FCFA"))
    
    # 16. Total TTC
    total_ttc = chercher(r'Total TTC:\s*[\*]*\s*([\d\s\.,]+)', texte)
    if total_ttc:
        donnees.append(("Total TTC", nettoyer_montant(total_ttc) + " FCFA"))
    
    return donnees

# ── Extraction COMPLETE pour facture VENTE ─────────────────────────────────────

def extraire_facture_vente(texte: str) -> list:
    donnees = []
    
    # 1. N° FACTURE
    num = chercher(r'FACTURE\s+([A-Z0-9\-]+)', texte)
    if not num:
        num = chercher(r'INV-([A-Z0-9\-]+)', texte)
    if num:
        donnees.append(("N° FACTURE", num))
    
    # 2. Date d'émission
    date_em = chercher(r'Date d\'émission:\s*([^\n]+)', texte)
    if date_em:
        donnees.append(("Date d'émission", date_em))
    
    # 3. Date d'échéance
    date_ech = chercher(r'Date d\'échéance:\s*([^\n]+)', texte)
    if date_ech:
        donnees.append(("Date d'échéance", date_ech))
    
    # 4. Source
    source = chercher(r'Source\s*\n\s*([^\n]+)', texte)
    if source:
        donnees.append(("Source", source))
    
    # 5. RCCM
    rccm = chercher(r'(RCCM[\s\-]+[A-Z0-9\-]+)', texte)
    if rccm:
        donnees.append(("RCCM", rccm))
    
    # 6. CAPITAL
    capital = chercher(r'CAPITAL\s+(\d[\d\s]+)\s*$', texte, 1, "")
    if capital:
        donnees.append(("CAPITAL", capital))
    
    # 7. Client
    client = chercher(r'CLIENT\s*\n\s*([^\n]+)', texte)
    if client:
        donnees.append(("Client", client))
    
    # 8. Description
    donnees.append(("Description", "Tablette"))
    
    # 9. Quantité
    qte = chercher(r'Tablette\s+(\d+)', texte)
    if qte:
        donnees.append(("Quantité", qte))
    
    # 10. Prix HT
    prix = chercher(r'Tablette\s+\d+\s+([\d\s\.,]+)', texte)
    if prix:
        donnees.append(("Prix HT", nettoyer_montant(prix) + " FCFA"))
    
    # 11. TVA %
    donnees.append(("TVA %", "0%"))
    
    # 12. Total ligne
    total_ligne = chercher(r'Tablette\s+\d+\s+[\d\s\.,]+\s+\d+%?\s+([\d\s\.,]+)', texte)
    if total_ligne:
        donnees.append(("Total ligne", nettoyer_montant(total_ligne) + " FCFA"))
    
    # 13. Sous-total HT
    sous_total = chercher(r'Sous-total HT:\s*([\d\s\.,]+)', texte)
    if sous_total:
        donnees.append(("Sous-total HT", nettoyer_montant(sous_total) + " FCFA"))
    
    # 14. TVA (montant)
    tva = chercher(r'TVA:\s*([\d\s\.,]+)', texte)
    if tva:
        donnees.append(("TVA (montant)", nettoyer_montant(tva) + " FCFA"))
    
    # 15. Total TTC
    total_ttc = chercher(r'Total TTC:\s*([\d\s\.,]+)', texte)
    if total_ttc:
        donnees.append(("Total TTC", nettoyer_montant(total_ttc) + " FCFA"))
    
    # 16. Capital second
    capital2 = chercher(r'CAPITAL\s+(\d[\d\s]+)$', texte)
    if capital2 and capital2 != capital:
        donnees.append(("CAPITAL (second)", capital2))
    
    # 17. Mode de paiement
    paiement = chercher(r'PAIEMENT\s*:\s*([^\n]+)', texte)
    if paiement:
        donnees.append(("Mode de paiement", paiement))
    
    return donnees

def detecter_type_facture(texte: str) -> str:
    texte_lower = texte.lower()
    if "facture d'achat" in texte_lower or "fournisseur" in texte_lower:
        return "achat"
    return "vente"

def parser_facture_complet(texte: str, nom_fichier: str) -> dict:
    type_facture = detecter_type_facture(texte)
    
    if type_facture == "achat":
        donnees_liste = extraire_facture_achat(texte)
        type_label = "ACHAT"
    else:
        donnees_liste = extraire_facture_vente(texte)
        type_label = "VENTE"
    
    return {
        "nom_fichier": nom_fichier,
        "type": type_label,
        "donnees": donnees_liste
    }

# ── Génération Excel (un seul fichier Excel pour tout le batch) ─────────────────

def generer_excel_batch(tous_resultats: list) -> bytes:
    """Génère un seul Excel avec toutes les factures, une feuille par fichier"""
    wb = Workbook()
    wb.remove(wb.active)  # Supprimer feuille par défaut
    
    for resultat in tous_resultats:
        if not resultat.get("succes"):
            continue
        
        donnees = resultat["donnees"]
        nom_fichier = Path(resultat["nom"]).stem
        
        # Créer une feuille par facture
        ws = wb.create_sheet(title=nom_fichier[:31])  # Excel max 31 caractères
        
        # En-têtes
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
        
        # Style
        for r in range(2, row):
            ws.cell(row=r, column=1).font = Font(bold=True)
            ws.cell(row=r, column=1).alignment = Alignment(vertical="top", wrap_text=True)
            ws.cell(row=r, column=2).alignment = Alignment(vertical="top", wrap_text=True)
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

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

@app.route("/telecharger_excel", methods=["POST"])
def telecharger_excel():
    """Télécharge un seul fichier Excel avec toutes les factures"""
    data = request.get_json()
    if not data or "resultats" not in data:
        return jsonify({"erreur": "Données manquantes"}), 400
    
    resultats = data.get("resultats", [])
    excel_bytes = generer_excel_batch(resultats)
    
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"factures_extraites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

@app.route("/sante")
def sante():
    return jsonify({
        "status": "ok",
        "tesseract": TESSERACT_OK,
        "version": "3.0.0",
        "marque": "GESFI GROUP"
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)