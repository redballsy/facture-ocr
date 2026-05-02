import sys
import os
import re
import io
import json
from pathlib import Path
from datetime import datetime

from flask import Flask, render_template, request, jsonify, send_file

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
    print("Lance: pip install flask pytesseract pillow opencv-python openpyxl")
    sys.exit(1)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB max

EXTENSIONS = {"jpg", "jpeg", "png", "bmp", "tiff", "tif", "webp", "gif"}

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

COULEURS = {
    "header_bg": "1B2A4A", "header_font": "FFFFFF",
    "titre_bg": "2E4172", "titre_font": "FFFFFF",
    "ligne_paire": "EEF2F7", "total_bg": "1B2A4A",
    "total_font": "FFFFFF", "sous_total": "D6E0F0", "border": "AABBD0",
}

# ── OCR ───────────────────────────────────────────────────────────────────────

def pretraiter_image(img_bytes: bytes) -> Image.Image:
    """Prétraite l'image pour améliorer l'OCR"""
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

# ── Parsing amélioré ─────────────────────────────────────────────────────────

def net(s):
    return s.strip() if s else ""

def chercher(pattern, texte, groupe=1, defaut=""):
    m = re.search(pattern, texte, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return net(m.group(groupe)) if m else defaut

def nettoyer_montant(montant_str):
    """Nettoie une chaîne de caractères pour extraire un montant numérique"""
    if not montant_str:
        return ""
    # Enlève les espaces et remplace les virgules par des points
    montant = re.sub(r'[^\d,\.]', '', montant_str)
    montant = montant.replace(',', '.')
    # Ne garde qu'un seul point décimal
    parties = montant.split('.')
    if len(parties) > 2:
        montant = parties[0] + '.' + ''.join(parties[1:])
    return montant

def parser_facture(texte: str) -> dict:
    lignes = [l.strip() for l in texte.split("\n") if l.strip()]
    donnees = {}
    
    # Extraction du numéro de facture
    donnees["numero_facture"] = chercher(
        r'(?:N°|Numéro|Facture)[\s:]*([A-Z0-9\-]+)', texte
    )
    if not donnees["numero_facture"]:
        donnees["numero_facture"] = chercher(r'\b(?:INV|FAC)[-\w]+\b', texte)
    
    # Extraction des dates
    dates = re.findall(r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})', texte)
    donnees["date_emission"] = dates[0] if len(dates) > 0 else ""
    donnees["date_echeance"] = dates[1] if len(dates) > 1 else ""
    
    # Extraction des informations fournisseur
    fournisseur_nom = chercher(r'^([A-Z]{2,}.*?)$', texte, 1, "")
    if not fournisseur_nom:
        fournisseur_nom = "GESFI"
    
    donnees["fournisseur"] = {
        "nom": fournisseur_nom,
        "adresse": chercher(r'(\d+\s+BP\s+[A-Z]+,\s*[A-Z]+,\s*Côte d\'Ivoire)', texte),
        "rccm": chercher(r'(RCCM[\s\-]+[A-Z0-9\-]+)', texte),
        "capital": chercher(r'CAPITAL\s+(\d[\d\s]+)', texte),
        "telephone": "",
        "email": "",
    }
    
    # Extraction du client
    client_nom = chercher(r'CLIENT\s+([^\n]+)', texte)
    if not client_nom:
        client_nom = chercher(r'Client[:\s]+([^\n]+)', texte)
    donnees["client"] = {"nom": client_nom, "adresse": "", "telephone": ""}
    
    # Extraction des lignes de facture
    resultats = []
    
    # Pattern spécifique pour la facture
    # Cherche "Tablette" suivi des montants
    pattern_tablette = r'Tablette\s+(\d+)\s+([\d\s,.]+)\s+(\d+%?)\s+([\d\s,.]+)'
    m = re.search(pattern_tablette, texte, re.IGNORECASE)
    if m:
        resultats.append({
            "description": "Tablette",
            "quantite": net(m.group(1)),
            "prix_unitaire_ht": net(m.group(2)),
            "tva_pct": net(m.group(3)),
            "total": net(m.group(4)),
        })
    else:
        # Pattern générique pour les lignes de facture
        pattern_ligne = re.compile(
            r'([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-/]{2,})\s+'
            r'(\d+)\s+'
            r'([\d\s,.]+)\s+'
            r'(\d+%?)\s*'
            r'(?:[A-Z]+)?\s*'
            r'([\d\s,.]+)'
        )
        for ligne in lignes:
            m = pattern_ligne.search(ligne)
            if m:
                resultats.append({
                    "description": net(m.group(1)),
                    "quantite": net(m.group(2)),
                    "prix_unitaire_ht": net(m.group(3)),
                    "tva_pct": net(m.group(4)),
                    "total": net(m.group(5)),
                })
    
    donnees["lignes"] = resultats
    
    # Extraction des totaux - améliorée
    sous_total = chercher(r'Sous-total HT:\s*([\d\s,.]+)', texte)
    if not sous_total:
        sous_total = chercher(r'Sous-total\s+HT\s+([\d\s,.]+)', texte)
    donnees["sous_total_ht"] = sous_total
    
    tva_montant = chercher(r'TVA:\s*([\d\s,.]+)', texte)
    if not tva_montant:
        tva_montant = chercher(r'TVA\s+([\d\s,.]+)', texte)
    donnees["tva_montant"] = tva_montant
    
    total_ttc = chercher(r'Total TTC:\s*([\d\s,.]+)', texte)
    if not total_ttc:
        total_ttc = chercher(r'Total\s+TTC\s+([\d\s,.]+)', texte)
    if not total_ttc:
        total_ttc = chercher(r'TOTAL TTC\s+([\d\s,.]+)', texte)
    donnees["total_ttc"] = total_ttc
    
    # Extraction du mode de paiement
    mode_paiement = chercher(r'PAIEMENT WAGE\s*:\s*\+225\s+([\d\s]+)', texte)
    if not mode_paiement:
        mode_paiement = chercher(r'PAIEMENT\s*:\s*([^\n]+)', texte)
    donnees["mode_paiement"] = f"WAVE: +225 {mode_paiement}" if mode_paiement else ""
    
    # Devise
    donnees["devise"] = "FCFA"
    donnees["notes"] = ""
    donnees["texte_brut"] = texte
    
    # Nettoyage des montants
    for champ in ['sous_total_ht', 'tva_montant', 'total_ttc']:
        if donnees.get(champ):
            donnees[champ] = nettoyer_montant(donnees[champ])
    
    for ligne in donnees.get('lignes', []):
        for champ in ['prix_unitaire_ht', 'total']:
            if ligne.get(champ):
                ligne[champ] = nettoyer_montant(ligne[champ])
        if ligne.get('tva_pct'):
            ligne['tva_pct'] = re.sub(r'[^\d%]', '', ligne['tva_pct'])
    
    return donnees

# ── Excel ─────────────────────────────────────────────────────────────────────

def v(val, defaut=""):
    return val if val else defaut

def style_cell(cell, bold=False, bg=None, font_color="000000",
               align="left", wrap=False, border=False, size=10):
    cell.font = Font(bold=bold, color=font_color, name="Arial", size=size)
    if bg:
        cell.fill = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        s = Side(style="thin", color=COULEURS["border"])
        cell.border = Border(left=s, right=s, top=s, bottom=s)

def generer_excel(donnees: dict, nom_fichier: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Facture"
    r = 1

    for col, w in {"A":30, "B":20, "C":20, "D":12, "E":12, "F":18}.items():
        ws.column_dimensions[col].width = w

    # Titre
    ws.merge_cells(f"A{r}:F{r}")
    ws[f"A{r}"].value = f"FACTURE — {v(donnees.get('numero_facture'), 'GESFI')}"
    style_cell(ws[f"A{r}"], bold=True, bg=COULEURS["titre_bg"],
               font_color=COULEURS["titre_font"], align="center", size=12)
    ws.row_dimensions[r].height = 22
    r += 1

    # En-tête
    infos = [
        ("N° Facture", v(donnees.get("numero_facture"))),
        ("Date d'émission", v(donnees.get("date_emission"))),
        ("Date d'échéance", v(donnees.get("date_echeance"))),
        ("Mode de paiement", v(donnees.get("mode_paiement"))),
    ]

    for label, valeur in infos:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = valeur
        ws.merge_cells(f"B{r}:F{r}")
        style_cell(ws[f"A{r}"], bold=True, bg=COULEURS["sous_total"], border=True)
        style_cell(ws[f"B{r}"], wrap=True, border=True)
        ws.row_dimensions[r].height = 16
        r += 1

    r += 1

    # FOURNISSEUR
    ws.merge_cells(f"A{r}:F{r}")
    ws[f"A{r}"].value = "FOURNISSEUR"
    style_cell(ws[f"A{r}"], bold=True, bg=COULEURS["header_bg"],
               font_color=COULEURS["header_font"], align="center")
    r += 1

    f = donnees.get("fournisseur", {}) or {}
    infos_fournisseur = [
        ("Nom", v(f.get("nom"))),
        ("Adresse", v(f.get("adresse"))),
        ("RCCM", v(f.get("rccm"))),
        ("Capital", v(f.get("capital"))),
    ]

    for label, valeur in infos_fournisseur:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = valeur
        ws.merge_cells(f"B{r}:F{r}")
        style_cell(ws[f"A{r}"], bold=True, border=True)
        style_cell(ws[f"B{r}"], wrap=True, border=True)
        r += 1

    r += 1

    # CLIENT
    ws.merge_cells(f"A{r}:F{r}")
    ws[f"A{r}"].value = "CLIENT"
    style_cell(ws[f"A{r}"], bold=True, bg=COULEURS["header_bg"],
               font_color=COULEURS["header_font"], align="center")
    r += 1

    c = donnees.get("client", {}) or {}
    ws[f"A{r}"] = "Nom"
    ws[f"B{r}"] = v(c.get("nom"))
    ws.merge_cells(f"B{r}:F{r}")
    style_cell(ws[f"A{r}"], bold=True, border=True)
    style_cell(ws[f"B{r}"], wrap=True, border=True)
    r += 2

    # Tableau des lignes
    for i, titre in enumerate(["Description", "Quantité", "Prix HT", "TVA %", "Total"], 1):
        cell = ws[f"{get_column_letter(i)}{r}"]
        cell.value = titre
        style_cell(cell, bold=True, bg=COULEURS["header_bg"],
                   font_color=COULEURS["header_font"], align="center", border=True)
    ws.row_dimensions[r].height = 18
    r += 1

    lignes_fact = donnees.get("lignes", []) or []
    if lignes_fact:
        for i, ligne in enumerate(lignes_fact):
            bg = COULEURS["ligne_paire"] if i % 2 == 0 else None
            vals = [
                v(ligne.get("description")),
                v(ligne.get("quantite")),
                v(ligne.get("prix_unitaire_ht")),
                v(ligne.get("tva_pct")),
                v(ligne.get("total"))
            ]
            for j, (val_, al) in enumerate(zip(vals, ["left", "center", "right", "center", "right"]), 1):
                cell = ws[f"{get_column_letter(j)}{r}"]
                cell.value = val_
                style_cell(cell, bg=bg, align=al, border=True, wrap=True)
            ws.row_dimensions[r].height = 16
            r += 1
    else:
        ws.merge_cells(f"A{r}:E{r}")
        ws[f"A{r}"].value = "Aucune ligne détectée"
        style_cell(ws[f"A{r}"], font_color="AA0000", border=True)
        r += 1

    r += 1

    # Totaux
    devise = v(donnees.get("devise"), "FCFA")
    for libelle, montant, bg_t, is_tot in [
        ("Sous-total HT", v(donnees.get("sous_total_ht")), COULEURS["sous_total"], False),
        ("TVA", v(donnees.get("tva_montant")), COULEURS["sous_total"], False),
        ("TOTAL TTC", v(donnees.get("total_ttc")), COULEURS["total_bg"], True),
    ]:
        ws.merge_cells(f"A{r}:C{r}")
        ws[f"A{r}"].value = libelle
        fc = COULEURS["total_font"] if is_tot else "000000"
        style_cell(ws[f"A{r}"], bold=is_tot, bg=bg_t, font_color=fc,
                   align="right", border=True, size=11 if is_tot else 10)
        ws.merge_cells(f"D{r}:E{r}")
        ws[f"D{r}"].value = f"{montant} {devise}".strip() if montant else ""
        style_cell(ws[f"D{r}"], bold=is_tot, bg=bg_t, font_color=fc,
                   align="right", border=True, size=11 if is_tot else 10)
        ws.row_dimensions[r].height = 20 if is_tot else 16
        r += 1

    # Onglet texte brut
    ws_raw = wb.create_sheet("Texte OCR brut")
    ws_raw["A1"] = "TEXTE BRUT OCR (vérification)"
    ws_raw["A1"].font = Font(bold=True, color="AA0000")
    ws_raw["A2"] = donnees.get("texte_brut", "")
    ws_raw.column_dimensions["A"].width = 80

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Routes Flask ──────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html", tesseract_ok=TESSERACT_OK)

@app.route("/analyser", methods=["POST"])
def analyser():
    if not TESSERACT_OK:
        return jsonify({"erreur": "Tesseract non installé sur le serveur."}), 500

    if "fichier" not in request.files:
        return jsonify({"erreur": "Aucun fichier reçu."}), 400

    fichier = request.files["fichier"]
    if not fichier.filename:
        return jsonify({"erreur": "Nom de fichier vide."}), 400

    ext = fichier.filename.rsplit(".", 1)[-1].lower()
    if ext not in EXTENSIONS:
        return jsonify({"erreur": f"Format non supporté: .{ext}"}), 400

    img_bytes = fichier.read()
    texte = extraire_texte(img_bytes)

    if not texte.strip():
        return jsonify({"erreur": "Aucun texte détecté dans l'image."}), 422

    donnees = parser_facture(texte)
    donnees["nom_fichier"] = fichier.filename
    donnees["mots_extraits"] = len(texte.split())
    return jsonify(donnees)

@app.route("/telecharger", methods=["POST"])
def telecharger():
    donnees = request.get_json()
    if not donnees:
        return jsonify({"erreur": "Données manquantes."}), 400

    nom_fichier = donnees.get("nom_fichier", "facture")
    excel_bytes = generer_excel(donnees, nom_fichier)
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
        "timestamp": datetime.now().isoformat()
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)