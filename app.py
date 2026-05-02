import sys
import os
import re
import io
import json
import zipfile
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, render_template, request, jsonify, send_file, after_this_request

try:
    import pytesseract
    from PIL import Image
    import cv2
    import numpy as np
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"❌ Dépendance manquante: {e}")
    sys.exit(1)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB max pour 100 fichiers
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "gesfi-secret-key-2026")

EXTENSIONS = {"jpg", "jpeg", "png", "bmp", "tiff", "tif", "webp", "gif"}

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

# Couleurs pour Excel
COULEURS_EXCEL = {
    "header_bg": COULEURS_GESFI["primary"],
    "header_font": COULEURS_GESFI["text_light"],
    "titre_bg": COULEURS_GESFI["secondary"],
    "titre_font": COULEURS_GESFI["text_light"],
    "ligne_paire": "#F0F4F8",
    "total_bg": COULEURS_GESFI["primary"],
    "total_font": COULEURS_GESFI["text_light"],
    "sous_total": "#E8F0F8",
    "border": COULEURS_GESFI["border"],
}

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

def extraire_texte(img_bytes: bytes) -> str:
    img = pretraiter_image(img_bytes)
    config = "--oem 3 --psm 6"
    try:
        return pytesseract.image_to_string(img, lang="fra+eng", config=config)
    except Exception:
        try:
            return pytesseract.image_to_string(img, lang="eng", config=config)
        except Exception:
            return ""

# ── Détection du type de facture ─────────────────────────────────────────────

def detecter_type_facture(texte: str) -> str:
    """Détecte si c'est une facture de vente ou d'achat"""
    texte_lower = texte.lower()
    
    # Indicateurs facture d'achat
    if "facture d'achat" in texte_lower or "po-" in texte_lower or "livraison prévue" in texte_lower:
        return "achat"
    
    # Indicateurs facture de vente
    if "client" in texte_lower and "tablette" in texte_lower:
        return "vente"
    
    # Indicateurs facture standard
    if "fournisseur" in texte_lower and "client" in texte_lower:
        return "standard"
    
    return "standard"

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

# ── Extraction spécifique pour facture de VENTE (facture.png) ─────────────────

def extraire_facture_vente(texte: str) -> dict:
    """Extraction pour le template facture de vente"""
    donnees = {
        "type": "VENTE",
        "champs": {},
        "lignes": [],
        "autres": []
    }
    
    # Numéro de facture
    num_facture = chercher(r'(?:INV|N°|Numéro)[\s\-:]*([A-Z0-9\-]+)', texte)
    if not num_facture:
        num_facture = chercher(r'FACTURE[\s\w]*([A-Z0-9\-]+)', texte)
    donnees["champs"]["N° FACTURE"] = num_facture
    
    # Dates
    donnees["champs"]["Date d'émission"] = chercher(r'Date d\'émission:\s*([^\n]+)', texte)
    donnees["champs"]["Date d'échéance"] = chercher(r'Date d\'échéance:\s*([^\n]+)', texte)
    
    # Source / Adresse
    source = chercher(r'Source\s*([^\n]+)', texte)
    if not source:
        source = chercher(r'(\d+\s+BP\s+[A-Z]+,\s*[A-Z]+,\s*Côte', texte)
    donnees["champs"]["Source"] = source
    
    # RCCM et Capital
    donnees["champs"]["RCCM"] = chercher(r'(RCCM[\s\-]+[A-Z0-9\-]+)', texte)
    donnees["champs"]["CAPITAL"] = chercher(r'CAPITAL\s+(\d[\d\s]+)', texte)
    
    # Client
    client = chercher(r'CLIENT\s+([^\n]+)', texte)
    if not client:
        client = chercher(r'Client[:\s]+([^\n]+)', texte)
    donnees["champs"]["Client"] = client
    
    # Lignes de facture
    pattern_ligne = r'(Tablette|Imprimante|Produit|Article)[\s\n]+(\d+)[\s\n]+([\d\s\.,]+)[\s\n]+(\d+%?)[\s\n]+([\d\s\.,]+)'
    match = re.search(pattern_ligne, texte, re.IGNORECASE | re.DOTALL)
    if match:
        donnees["lignes"].append({
            "Description": net(match.group(1)),
            "Quantité": net(match.group(2)),
            "Prix HT": net(match.group(3)),
            "TVA %": net(match.group(4)),
            "Total": net(match.group(5))
        })
    
    # Totaux
    donnees["champs"]["Sous-total HT"] = chercher(r'Sous-total HT:\s*([\d\s\.,]+)', texte)
    donnees["champs"]["TVA"] = chercher(r'TVA:\s*([\d\s\.,]+)', texte)
    donnees["champs"]["Total TTC"] = chercher(r'Total TTC:\s*([\d\s\.,]+)', texte)
    
    # Paiement
    donnees["champs"]["Mode de paiement"] = chercher(r'PAIEMENT\s*:\s*([^\n]+)', texte)
    
    # Autres informations
    lignes_texte = [l.strip() for l in texte.split('\n') if l.strip()]
    champs_connus = set(donnees["champs"].keys()) | {"Description", "Quantité", "Prix HT", "TVA %", "Total"}
    
    for ligne in lignes_texte:
        if ligne and not any(c in ligne for c in champs_connus):
            if ligne not in donnees["autres"]:
                donnees["autres"].append(ligne)
    
    return donnees

# ── Extraction spécifique pour facture d'ACHAT (facture2.png) ─────────────────

def extraire_facture_achat(texte: str) -> dict:
    """Extraction pour le template facture d'achat"""
    donnees = {
        "type": "ACHAT",
        "champs": {},
        "lignes": [],
        "autres": []
    }
    
    # Numéro de facture
    donnees["champs"]["N° FACTURE"] = chercher(r'N°\s*([A-Z0-9\-]+)', texte)
    if not donnees["champs"]["N° FACTURE"]:
        donnees["champs"]["N° FACTURE"] = chercher(r'PO-([A-Z0-9\-]+)', texte)
    
    # Dates
    donnees["champs"]["Date"] = chercher(r'Date:\s*([^\n]+)', texte)
    donnees["champs"]["Livraison prévue"] = chercher(r'Livraison prévue:\s*([^\n]+)', texte)
    
    # Source
    donnees["champs"]["Source"] = chercher(r'Source\s*([^\n]+)', texte)
    
    # Fournisseur
    donnees["champs"]["Fournisseur Nom"] = chercher(r'FOURNISSEUR\s*([^\n]+)', texte)
    donnees["champs"]["Fournisseur Adresse"] = chercher(r'FOURNISSEUR\s*[^\n]+\s*([^\n]+)', texte, 1, "")
    donnees["champs"]["Fournisseur Téléphone"] = chercher(r'Tél:\s*([^\n]+)', texte)
    donnees["champs"]["Fournisseur Email"] = chercher(r'Email:\s*([^\n]+)', texte)
    
    # Lignes de facture
    pattern_ligne = r'(Imprimante de bureau|Tablette|Produit)[\s\n]+([\d\.,]+)[\s\n]+([\d\s\.,]+)[\s\n]+(\d+%?)[\s\n]+([\d\s\.,]+)'
    match = re.search(pattern_ligne, texte, re.IGNORECASE | re.DOTALL)
    if match:
        donnees["lignes"].append({
            "Description": net(match.group(1)),
            "Quantité": net(match.group(2)),
            "Prix Unitaire": net(match.group(3)),
            "TVA %": net(match.group(4)),
            "Total HT": net(match.group(5))
        })
    
    # Totaux
    donnees["champs"]["Sous-total HT"] = chercher(r'Sous-total HT:\s*([\d\s\.,]+)', texte)
    donnees["champs"]["TVA"] = chercher(r'TVA:\s*([\d\s\.,]+)', texte)
    donnees["champs"]["Total TTC"] = chercher(r'Total TTC:\s*([\d\s\.,]+)', texte)
    
    # Autres informations
    lignes_texte = [l.strip() for l in texte.split('\n') if l.strip()]
    champs_connus = set(donnees["champs"].keys())
    
    for ligne in lignes_texte:
        if ligne and not any(c in ligne for c in champs_connus):
            if ligne not in donnees["autres"]:
                donnees["autres"].append(ligne)
    
    return donnees

# ── Extraction générique ─────────────────────────────────────────────────────

def extraire_facture_generique(texte: str) -> dict:
    """Extraction générique pour tout type de facture"""
    donnees = {
        "type": "STANDARD",
        "champs": {},
        "lignes": [],
        "autres": []
    }
    
    # Extraction des paires clé:valeur
    patterns = [
        (r'(N°|Numéro|Facture)[\s:]+([^\n]+)', "N° FACTURE"),
        (r'Date[\s:]+([^\n]+)', "Date"),
        (r'Client[\s:]+([^\n]+)', "Client"),
        (r'Fournisseur[\s:]+([^\n]+)', "Fournisseur"),
        (r'Total[\s:]+([^\n]+)', "Total"),
    ]
    
    for pattern, nom in patterns:
        valeur = chercher(pattern, texte, 2)
        if valeur:
            donnees["champs"][nom] = valeur
    
    return donnees

# ── Parser principal ─────────────────────────────────────────────────────────

def parser_facture_complet(texte: str) -> dict:
    """Parse toute facture selon son type détecté"""
    type_facture = detecter_type_facture(texte)
    
    if type_facture == "vente":
        donnees = extraire_facture_vente(texte)
    elif type_facture == "achat":
        donnees = extraire_facture_achat(texte)
    else:
        donnees = extraire_facture_generique(texte)
    
    donnees["texte_brut"] = texte
    donnees["type_detecte"] = type_facture
    donnees["devise"] = "FCFA"
    
    # Nettoyage des montants
    for champ, valeur in donnees["champs"].items():
        if any(mot in champ for mot in ["HT", "TTC", "TVA", "Prix", "Total", "Sous-total"]):
            donnees["champs"][champ] = nettoyer_montant(valeur)
    
    return donnees

# ── Génération Excel avec structure demandée ─────────────────────────────────

def generer_excel_structure(donnees: dict, nom_fichier: str) -> bytes:
    """Génère Excel avec colonne CHAMP | VALEUR | AUTRES"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Facture"
    
    # En-têtes
    headers = ["CHAMP", "VALEUR", "AUTRES INFORMATIONS"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color=COULEURS_EXCEL["header_font"], size=11)
        cell.fill = PatternFill("solid", start_color=COULEURS_EXCEL["header_bg"].replace("#", ""))
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 50
    
    row = 2
    
    # 1. Type de facture
    ws.cell(row=row, column=1, value="TYPE FACTURE")
    ws.cell(row=row, column=2, value=donnees.get("type", "STANDARD"))
    row += 1
    
    # 2. Champs extraits
    for champ, valeur in donnees.get("champs", {}).items():
        if valeur:
            ws.cell(row=row, column=1, value=champ)
            ws.cell(row=row, column=2, value=str(valeur))
            row += 1
    
    # 3. Lignes de facture
    for ligne in donnees.get("lignes", []):
        for champ, valeur in ligne.items():
            ws.cell(row=row, column=1, value=f"  └ {champ}")
            ws.cell(row=row, column=2, value=str(valeur))
            row += 1
    
    # 4. Autres informations
    for autre in donnees.get("autres", []):
        if autre and len(autre) > 2:
            ws.cell(row=row, column=1, value="Autre")
            ws.cell(row=row, column=3, value=str(autre))
            row += 1
    
    # Style des cellules
    for r in range(2, row):
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if c == 1:
                cell.font = Font(bold=True)
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Traitement batch de plusieurs fichiers ───────────────────────────────────

def traiter_un_fichier(fichier_info: tuple) -> dict:
    """Traite un seul fichier et retourne les données"""
    idx, fichier = fichier_info
    try:
        img_bytes = fichier.read()
        fichier.seek(0)  # Reset pour lecture ultérieure
        
        texte = extraire_texte(img_bytes)
        if not texte.strip():
            return {"index": idx, "nom": fichier.filename, "erreur": "Aucun texte détecté"}
        
        donnees = parser_facture_complet(texte)
        donnees["nom_fichier"] = fichier.filename
        
        return {"index": idx, "nom": fichier.filename, "succes": True, "donnees": donnees}
    except Exception as e:
        return {"index": idx, "nom": fichier.filename, "erreur": str(e)}

# ── Routes Flask ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html", tesseract_ok=TESSERACT_OK, couleurs=COULEURS_GESFI)

@app.route("/analyser_batch", methods=["POST"])
def analyser_batch():
    """Analyse jusqu'à 100 fichiers en batch"""
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
    
    # Traitement parallèle
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(traiter_un_fichier, ft): ft for ft in fichiers_a_traiter}
        for future in as_completed(futures):
            resultats.append(future.result())
    
    # Trier par index
    resultats.sort(key=lambda x: x.get("index", 0))
    
    return jsonify({
        "total": len(resultats),
        "succes": len([r for r in resultats if r.get("succes")]),
        "erreurs": len([r for r in resultats if r.get("erreur")]),
        "resultats": resultats
    })

@app.route("/telecharger_batch", methods=["POST"])
def telecharger_batch():
    """Télécharge tous les fichiers Excel en un ZIP"""
    data = request.get_json()
    if not data or "resultats" not in data:
        return jsonify({"erreur": "Données manquantes"}), 400
    
    resultats = data.get("resultats", [])
    
    # Créer un fichier ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for resultat in resultats:
            if resultat.get("succes") and resultat.get("donnees"):
                donnees = resultat["donnees"]
                nom_original = Path(resultat["nom"]).stem
                excel_bytes = generer_excel_structure(donnees, nom_original)
                zip_file.writestr(f"{nom_original}_extrait.xlsx", excel_bytes)
    
    zip_buffer.seek(0)
    
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"factures_extraites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    )

@app.route("/telecharger_unique", methods=["POST"])
def telecharger_unique():
    """Télécharge un seul fichier Excel"""
    donnees = request.get_json()
    if not donnees:
        return jsonify({"erreur": "Données manquantes"}), 400
    
    nom_fichier = donnees.get("nom_fichier", "facture")
    excel_bytes = generer_excel_structure(donnees, nom_fichier)
    nom_sortie = f"facture_{Path(nom_fichier).stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
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
        "timestamp": datetime.now().isoformat()
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)