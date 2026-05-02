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

# ── Parsing ───────────────────────────────────────────────────────────────────

def net(s):
    return s.strip() if s else ""

def chercher(pattern, texte, groupe=1, defaut=""):
    m = re.search(pattern, texte, re.IGNORECASE | re.MULTILINE)
    return net(m.group(groupe)) if m else defaut

def parser_facture(texte: str) -> dict:
    lignes = [l.strip() for l in texte.split("\n") if l.strip()]
    donnees = {}

    donnees["numero_facture"] = chercher(
        r"(?:INV|FAC|N°?|Num[eé]ro?)[.\s\-:]*([A-Z0-9\-]+)", texte
    ) or chercher(r"(\bINV[-\w]+)", texte)

    dates = re.findall(r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})", texte)
    donnees["date_emission"] = dates[0] if len(dates) > 0 else ""
    donnees["date_echeance"] = dates[1] if len(dates) > 1 else ""

    fournisseur_nom = ""
    for ligne in lignes[:8]:
        if any(m in ligne.lower() for m in ["sarl","sa ","sas","eurl","ltd","inc","source","gesfi"]):
            fournisseur_nom = ligne
            break
    if not fournisseur_nom and lignes:
        fournisseur_nom = lignes[0]

    donnees["fournisseur"] = {
        "nom": fournisseur_nom,
        "adresse": chercher(r"(\d+\s+(?:BP|rue|avenue|bd|blvd|allée)[^\n|]+)", texte),
        "rccm": chercher(r"(RCCM\s+[A-Z0-9\-]+)", texte),
        "capital": chercher(r"[Cc]apital[^\d]*(\d[\d\s]+)", texte),
        "telephone": chercher(r"(?:Tél|Tel|Phone|Téléphone)[^\d]*(\+?[\d\s\-\.]{8,})", texte),
        "email": chercher(r"[\w.\-]+@[\w.\-]+\.[a-z]{2,}", texte, 0),
    }

    client_nom = chercher(r"(?:Client|Destinataire)[:\s]+([^\n|]{2,40})", texte)
    donnees["client"] = {"nom": client_nom, "adresse": "", "telephone": ""}

    resultats = []
    pattern_ligne = re.compile(
        r"([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\s\-/]{2,40})\s+"
        r"(\d+)\s+([\d\s,.']+)\s+(\d+\s*%?)\s+([\d\s,.']+)"
    )
    for ligne in lignes:
        m = pattern_ligne.search(ligne)
        if m:
            resultats.append({
                "description": net(m.group(1)), "quantite": net(m.group(2)),
                "prix_unitaire_ht": net(m.group(3)), "tva_pct": net(m.group(4)),
                "total": net(m.group(5)),
            })

    if not resultats:
        mots_exclus = ["total","sous","tva","date","description","client",
                       "facture","adresse","rccm","capital","paiement"]
        for ligne in lignes:
            if re.search(r"[A-Za-zÀ-ÿ]{3,}", ligne) and re.search(r"\d{3,}", ligne):
                if not any(m in ligne.lower() for m in mots_exclus):
                    nombres = re.findall(r"[\d\s,.']+", ligne)
                    desc = re.sub(r"[\d\s,.']+", "", ligne).strip()
                    if desc and nombres:
                        resultats.append({
                            "description": desc,
                            "quantite": net(nombres[0]) if len(nombres)>0 else "",
                            "prix_unitaire_ht": net(nombres[1]) if len(nombres)>1 else "",
                            "tva_pct": net(nombres[2]) if len(nombres)>2 else "",
                            "total": net(nombres[-1]),
                        })

    donnees["lignes"] = resultats
    donnees["sous_total_ht"] = chercher(r"[Ss]ous.?[Tt]otal[^\d]*(\d[\d\s,.']+)", texte)
    donnees["tva_montant"] = chercher(r"TVA[^\d]*(\d[\d\s,.']+)", texte)
    donnees["total_ttc"] = chercher(r"[Tt]otal\s*(?:TTC|ttc|général|dû)[^\d]*(\d[\d\s,.']+)", texte)

    if not donnees["total_ttc"]:
        montants = re.findall(r"(\d[\d\s]{3,}[.,]?\d*)\s*(?:FCFA|XOF|€|EUR|USD|\$)?", texte)
        if montants:
            donnees["total_ttc"] = montants[-1].strip()

    devise = "FCFA"
    for d in ["FCFA","XOF","EUR","€","USD","$"]:
        if d in texte:
            devise = d
            break
    donnees["devise"] = devise
    donnees["mode_paiement"] = chercher(r"(?:PAIEMENT|Règlement|Payable)[^\n:]*[:\s]+([^\n]+)", texte)
    donnees["notes"] = ""
    donnees["texte_brut"] = texte
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

    for col, w in {"A":30,"B":20,"C":20,"D":12,"E":12,"F":18}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells(f"A{r}:F{r}")
    ws[f"A{r}"].value = f"FACTURE  {v(donnees.get('numero_facture'),'N/A')}  —  {v(donnees.get('fournisseur',{}).get('nom'),'Inconnu')}"
    style_cell(ws[f"A{r}"], bold=True, bg=COULEURS["titre_bg"],
               font_color=COULEURS["titre_font"], align="center", size=12)
    ws.row_dimensions[r].height = 22
    r += 1

    f = donnees.get("fournisseur", {}) or {}
    c = donnees.get("client", {}) or {}

    infos = [
        ("Fichier source", nom_fichier),
        ("N° Facture", v(donnees.get("numero_facture"))),
        ("Date d'émission", v(donnees.get("date_emission"))),
        ("Date d'échéance", v(donnees.get("date_echeance"))),
        ("Mode de paiement", v(donnees.get("mode_paiement"))),
        ("── FOURNISSEUR ──", ""),
        ("Nom", v(f.get("nom"))), ("Adresse", v(f.get("adresse"))),
        ("RCCM", v(f.get("rccm"))), ("Capital", v(f.get("capital"))),
        ("Téléphone", v(f.get("telephone"))), ("Email", v(f.get("email"))),
        ("── CLIENT ──", ""),
        ("Nom", v(c.get("nom"))), ("Adresse", v(c.get("adresse"))),
    ]

    for label, valeur in infos:
        ws[f"A{r}"] = label
        ws[f"B{r}"] = valeur
        ws.merge_cells(f"B{r}:F{r}")
        is_sec = label.startswith("──")
        style_cell(ws[f"A{r}"], bold=True, bg=COULEURS["sous_total"] if is_sec else None, border=True)
        style_cell(ws[f"B{r}"], bold=is_sec, wrap=True, border=True)
        ws.row_dimensions[r].height = 16
        r += 1

    r += 1

    for i, titre in enumerate(["Description","Quantité","Prix HT","TVA %","Montant TVA","Total"], 1):
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
            vals = [v(ligne.get("description")), v(ligne.get("quantite")),
                    v(ligne.get("prix_unitaire_ht")), v(ligne.get("tva_pct")),
                    "", v(ligne.get("total"))]
            for j, (val_, al) in enumerate(zip(vals,["left","center","right","center","right","right"]), 1):
                cell = ws[f"{get_column_letter(j)}{r}"]
                cell.value = val_
                style_cell(cell, bg=bg, align=al, border=True, wrap=True)
            ws.row_dimensions[r].height = 16
            r += 1
    else:
        ws.merge_cells(f"A{r}:F{r}")
        ws[f"A{r}"] = "⚠ Lignes non détectées automatiquement"
        style_cell(ws[f"A{r}"], font_color="AA0000", border=True)
        r += 1

    r += 1
    devise = v(donnees.get("devise"), "")
    for libelle, montant, bg_t, is_tot in [
        ("Sous-total HT", v(donnees.get("sous_total_ht")), COULEURS["sous_total"], False),
        ("TVA", v(donnees.get("tva_montant")), COULEURS["sous_total"], False),
        ("TOTAL TTC", v(donnees.get("total_ttc")), COULEURS["total_bg"], True),
    ]:
        ws.merge_cells(f"A{r}:D{r}")
        ws[f"A{r}"] = libelle
        fc = COULEURS["total_font"] if is_tot else "000000"
        style_cell(ws[f"A{r}"], bold=is_tot, bg=bg_t, font_color=fc,
                   align="right", border=True, size=11 if is_tot else 10)
        ws.merge_cells(f"E{r}:F{r}")
        ws[f"E{r}"] = f"{montant} {devise}".strip() if montant else ""
        style_cell(ws[f"E{r}"], bold=is_tot, bg=bg_t, font_color=fc,
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
    app.run(debug=True, port=5000)
