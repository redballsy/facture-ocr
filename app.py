import sys
import os
import re
import io
import json
import base64
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

from flask import Flask, render_template, request, jsonify, send_file

try:
    import pytesseract
    from PIL import Image, ImageEnhance, ImageFilter, ImageDraw, ImageFont
    import numpy as np
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from pdf2image import convert_from_bytes
except ImportError as e:
    print(f"❌ Dépendance manquante: {e}")
    sys.exit(1)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "gesfi-secret-key-2026")

EXTENSIONS = {"png", "pdf"}  # PNG et PDF uniquement

COULEURS_GESFI = {"primary": "#1B3A5C", "secondary": "#2E5A88", "accent": "#D4AF37"}
COULEURS_EXCEL = {"header_bg": "1B3A5C", "header_font": "FFFFFF"}

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

def extension_autorisee(filename: str) -> bool:
    if '.' not in filename:
        return False
    return filename.rsplit('.', 1)[1].lower() in EXTENSIONS

def pretraiter_image(img_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(img_bytes))
    if img.mode != 'RGB':
        img = img.convert('RGB')
    if img.width < 1200:
        ratio = 1200 / img.width
        img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.LANCZOS)
    img = img.convert('L')
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = img.filter(ImageFilter.SHARPEN)
    return img

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
    try:
        images = convert_from_bytes(img_bytes, dpi=300)
        texte_complet = ""
        for i, image in enumerate(images):
            buf = io.BytesIO()
            image.save(buf, format='PNG')
            texte_complet += f"\n--- Page {i+1} ---\n{extraire_texte_image(buf.getvalue())}\n"
        return texte_complet
    except Exception as e:
        return ""

def extraire_texte(img_bytes: bytes, filename: str) -> str:
    ext = filename.lower().rsplit('.', 1)[-1] if '.' in filename else ''
    return traiter_pdf(img_bytes) if ext == 'pdf' else extraire_texte_image(img_bytes)

def net(s): return s.strip() if s else ""

def chercher(pattern, texte, groupe=1, defaut=""):
    m = re.search(pattern, texte, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return net(m.group(groupe)) if m else defaut

def to_float(s):
    try:
        cleaned = re.sub(r'[^\d,\.]', '', (s or "").replace(' ', '')).replace(',', '.')
        return float(cleaned)
    except:
        return 0.0

def fmt_fcfa(val: float) -> str:
    return f"{val:,.0f} FCFA".replace(",", " ")

def calculer_totaux(prix_ht: float, quantite: float, tva_pct: float) -> dict:
    sous_total_ht = prix_ht * quantite
    montant_tva = sous_total_ht * (tva_pct / 100) if tva_pct > 0 else 0.0
    total_ttc = sous_total_ht + montant_tva
    return {
        "prix_ht": prix_ht, "quantite": quantite,
        "sous_total_ht": sous_total_ht, "tva_pct": tva_pct,
        "montant_tva": montant_tva, "total_ttc": total_ttc,
    }

def detecter_type_facture(texte: str) -> str:
    texte_lower = texte.lower()
    if "facture d'achat" in texte_lower or "fournisseur" in texte_lower:
        return "achat"
    return "vente"

def extraire_facture_vente(texte: str) -> dict:
    d = {}
    num = chercher(r'INV-([A-Z0-9\-]+)', texte)
    if not num:
        num = chercher(r'FACTURE\s+([A-Z0-9\-]+)', texte)
    d["num_facture"] = ("INV-" + num) if (num and not num.startswith("INV")) else num
    d["date_emission"] = chercher(r"Date d['\u2019][\u00e9e]mission\s*:\s*([^\n]+)", texte)
    d["date_echeance"] = chercher(r"Date d['\u2019][\u00e9e]ch[\u00e9e]ance\s*:\s*([^\n]+)", texte)
    d["source"] = chercher(r'Source\s*\n\s*([^\n]+)', texte)
    d["rccm"] = chercher(r'(RCCM\s+CI[\-\s]+[A-Z0-9\-\s]+)', texte).strip()
    cap = chercher(r'CAPITAL\s+([\d\s]+)', texte)
    d["capital"] = cap.replace(' ', '') if cap else ""
    d["client"] = chercher(r'CLIENT\s*\n\s*([^\n]+)', texte)
    d["paiement"] = chercher(r'PAIEMENT\s*(?:WAVE)?\s*[:\-]?\s*([^\n]+)', texte)

    desc_m = re.search(r'([A-Za-z\u00C0-\u00FF][A-Za-z\u00C0-\u00FF\s]+?)\s+(\d+)\s+([\d\s\.,]+)\s+\d+%', texte)
    if desc_m:
        d["description"] = desc_m.group(1).strip()
        d["quantite_str"] = desc_m.group(2)
        d["prix_ht_str"] = desc_m.group(3)
    else:
        d["description"] = "Article"; d["quantite_str"] = "1"; d["prix_ht_str"] = ""

    tva_s = chercher(r'(\d+)\s*%', texte)
    d["tva_pct"] = float(tva_s) if tva_s else 0.0
    d.update(calculer_totaux(to_float(d.get("prix_ht_str","")), to_float(d.get("quantite_str","1")) or 1.0, d["tva_pct"]))
    return d

def extraire_facture_achat(texte: str) -> dict:
    d = {}
    num = chercher(r'PO-([A-Z0-9\-]+)', texte)
    if not num:
        num = chercher(r'N[\u00b0\u00b0\s]+([A-Z0-9\-]+)', texte)
    d["num_facture"] = ("PO-" + num) if (num and not num.startswith("PO")) else num
    d["date"] = chercher(r'Date\s*:\s*([^\n]+)', texte)
    d["livraison"] = chercher(r'Livraison pr[\u00e9e]vue\s*:\s*([^\n]+)', texte)
    d["source"] = chercher(r'Source\s*\n\s*([^\n]+)', texte)
    d["fournisseur_nom"] = chercher(r'FOURNISSEUR\s*\n\s*([^\n]+)', texte)
    d["fournisseur_adresse"] = chercher(r'FOURNISSEUR\s*[^\n]+\n\s*([^\n]+)', texte)
    d["fournisseur_tel"] = chercher(r'T[\u00e9e]l\s*:\s*([^\n]+)', texte)
    d["fournisseur_email"] = chercher(r'Email\s*:\s*([^\n]+)', texte)

    desc_m = re.search(r'(Imprimante\s+de\s+bureau|[A-Za-z\u00C0-\u00FF][A-Za-z\u00C0-\u00FF\s]+?)\s+([\d,\.]+)\s+([\d\s\.,]+)', texte)
    if desc_m:
        d["description"] = desc_m.group(1).strip()
        d["quantite_str"] = desc_m.group(2)
        d["prix_ht_str"] = desc_m.group(3)
    else:
        d["description"] = "Article"; d["quantite_str"] = "1"; d["prix_ht_str"] = ""

    tva_s = chercher(r'(\d+)\s*%', texte)
    d["tva_pct"] = float(tva_s) if tva_s else 0.0
    d.update(calculer_totaux(to_float(d.get("prix_ht_str","")), to_float(d.get("quantite_str","1")) or 1.0, d["tva_pct"]))
    return d

def parser_facture_complet(texte: str, nom_fichier: str, solde: bool = False) -> dict:
    type_facture = detecter_type_facture(texte)
    if type_facture == "achat":
        donnees = extraire_facture_achat(texte)
        donnees["type"] = "ACHAT"
    else:
        donnees = extraire_facture_vente(texte)
        donnees["type"] = "VENTE"
    donnees["nom_fichier"] = nom_fichier
    donnees["solde"] = solde
    return donnees

# ── Génération image de synthèse ─────────────────────────────────────────────

def generer_image_facture(donnees: dict) -> str:
    W, H = 820, 960
    img = Image.new("RGB", (W, H), color=(245, 247, 250))
    draw = ImageDraw.Draw(img)

    C_DARK   = (27, 58, 92)
    C_ACCENT = (212, 175, 55)
    C_WHITE  = (255, 255, 255)
    C_LIGHT  = (235, 240, 248)
    C_GREEN  = (34, 139, 34)
    C_RED    = (180, 30, 30)
    C_GRAY   = (100, 110, 130)
    C_TEXT   = (30, 40, 60)
    C_SEC    = (46, 90, 136)

    def load_font(bold=False, size=13):
        try:
            name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
            return ImageFont.truetype(f"/usr/share/fonts/truetype/dejavu/{name}", size)
        except:
            return ImageFont.load_default()

    f_title  = load_font(True, 22)
    f_bold   = load_font(True, 14)
    f_reg    = load_font(False, 13)
    f_small  = load_font(False, 11)
    f_large  = load_font(True, 19)
    f_num    = load_font(False, 12)

    y = 0

    # En-tête
    draw.rectangle([(0, 0), (W, 72)], fill=C_DARK)
    draw.text((20, 12), "GESFI GROUP", font=f_title, fill=C_ACCENT)
    type_label = donnees.get("type", "FACTURE")
    num = donnees.get("num_facture", "—")
    draw.text((20, 42), f"FACTURE D'{type_label}   •   {num}", font=f_reg, fill=C_WHITE)

    solde = donnees.get("solde", False)
    badge_c = C_GREEN if solde else C_RED
    badge_t = "  SOLDE  " if solde else " NON SOLDE"
    draw.rectangle([(W - 145, 18), (W - 12, 54)], fill=badge_c)
    try:
        bbox = draw.textbbox((0,0), badge_t, font=f_bold)
        bw = bbox[2]-bbox[0]
    except:
        bw = len(badge_t)*8
    bx = W - 145 + (133 - bw)//2
    draw.text((bx, 28), badge_t, font=f_bold, fill=C_WHITE)
    y = 76

    row_h = 30

    def section(title, color=C_DARK):
        nonlocal y
        draw.rectangle([(0, y), (W, y+30)], fill=color)
        draw.text((14, y+8), title, font=f_bold, fill=C_WHITE)
        y += 30

    def row(label, value, alt=False):
        nonlocal y
        if not value:
            return
        draw.rectangle([(0, y), (W, y+row_h)], fill=(C_LIGHT if alt else C_WHITE))
        draw.text((14, y+8), label, font=f_bold, fill=C_GRAY)
        draw.text((300, y+8), str(value)[:60], font=f_reg, fill=C_TEXT)
        y += row_h

    def sep():
        nonlocal y
        draw.rectangle([(0, y), (W, y+4)], fill=C_ACCENT)
        y += 4

    # Infos générales
    section("  INFORMATIONS GENERALES")
    alt = False
    if type_label == "VENTE":
        for label, key in [
            ("Date d'emission", "date_emission"),
            ("Date d'echeance", "date_echeance"),
            ("Source", "source"),
            ("RCCM", "rccm"),
            ("Capital", "capital"),
        ]:
            row(label, donnees.get(key,""), alt); alt = not alt
    else:
        for label, key in [
            ("Date", "date"),
            ("Livraison prevue", "livraison"),
            ("Source", "source"),
        ]:
            row(label, donnees.get(key,""), alt); alt = not alt
    sep()

    # Interlocuteur
    if type_label == "VENTE":
        section("  CLIENT", C_SEC)
        row("Nom client", donnees.get("client",""), False)
        row("Mode de paiement", donnees.get("paiement",""), True)
    else:
        section("  FOURNISSEUR", C_SEC)
        alt2 = False
        for label, key in [
            ("Nom", "fournisseur_nom"),
            ("Adresse", "fournisseur_adresse"),
            ("Telephone", "fournisseur_tel"),
            ("Email", "fournisseur_email"),
        ]:
            row(label, donnees.get(key,""), alt2); alt2 = not alt2
    sep()

    # Detail article
    section("  DETAIL ARTICLE", C_SEC)
    row("Description", donnees.get("description","—"), False)
    row("Quantite", donnees.get("quantite_str","1"), True)
    row("Prix unitaire HT", fmt_fcfa(donnees.get("prix_ht", 0.0)), False)
    sep()

    # Calcul financier
    section("  CALCUL FINANCIER")
    tva_pct     = donnees.get("tva_pct", 0.0)
    sous_total  = donnees.get("sous_total_ht", 0.0)
    montant_tva = donnees.get("montant_tva", 0.0)
    total_ttc   = donnees.get("total_ttc", 0.0)
    qte_s       = donnees.get("quantite_str","1")
    prix_u      = donnees.get("prix_ht", 0.0)

    row(f"Sous-total HT  ({qte_s} x {fmt_fcfa(prix_u)})", fmt_fcfa(sous_total), False)
    if tva_pct > 0:
        row(f"TVA ({tva_pct:.0f}%)  = {fmt_fcfa(sous_total)} x {tva_pct:.0f}%", fmt_fcfa(montant_tva), True)
    else:
        row("TVA", "0% — exonere", True)

    # Total TTC mis en valeur
    draw.rectangle([(0, y), (W, y+46)], fill=C_ACCENT)
    draw.text((14, y+12), "TOTAL TTC", font=f_large, fill=C_DARK)
    ttc_str = fmt_fcfa(total_ttc)
    try:
        bbox = draw.textbbox((0,0), ttc_str, font=f_large)
        tw = bbox[2]-bbox[0]
    except:
        tw = len(ttc_str)*11
    draw.text((W - tw - 20, y+12), ttc_str, font=f_large, fill=C_DARK)
    y += 46
    sep()

    # Pied de page
    draw.rectangle([(0, H-32), (W, H)], fill=C_DARK)
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    draw.text((14, H-21), f"Genere le {ts} — GESFI GROUP v4.0", font=f_small, fill=C_ACCENT)
    nom = donnees.get("nom_fichier","")
    draw.text((W-len(nom)*6-10, H-21), nom, font=f_small, fill=C_WHITE)

    buf = io.BytesIO()
    img.save(buf, format="PNG", dpi=(150, 150))
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

# ── Génération Excel ─────────────────────────────────────────────────────────

def generer_excel_batch(tous_resultats: list) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    for resultat in tous_resultats:
        if not resultat.get("succes"):
            continue
        donnees = resultat["donnees"]
        ws = wb.create_sheet(title=Path(resultat["nom"]).stem[:31])

        for col, h in enumerate(["LIBELLE", "VALEUR"], 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill("solid", start_color="1B3A5C")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 50

        rows = [
            ("TYPE FACTURE", donnees.get("type","")),
            ("SOLDE", "OUI" if donnees.get("solde") else "NON"),
            ("N FACTURE", donnees.get("num_facture","")),
        ]
        if donnees.get("type") == "VENTE":
            rows += [
                ("Date emission", donnees.get("date_emission","")),
                ("Date echeance", donnees.get("date_echeance","")),
                ("Source", donnees.get("source","")),
                ("RCCM", donnees.get("rccm","")),
                ("Capital", donnees.get("capital","")),
                ("Client", donnees.get("client","")),
                ("Mode de paiement", donnees.get("paiement","")),
            ]
        else:
            rows += [
                ("Date", donnees.get("date","")),
                ("Livraison prevue", donnees.get("livraison","")),
                ("Source", donnees.get("source","")),
                ("Fournisseur Nom", donnees.get("fournisseur_nom","")),
                ("Fournisseur Adresse", donnees.get("fournisseur_adresse","")),
                ("Fournisseur Tel", donnees.get("fournisseur_tel","")),
                ("Fournisseur Email", donnees.get("fournisseur_email","")),
            ]
        rows += [
            ("Description", donnees.get("description","")),
            ("Quantite", donnees.get("quantite_str","")),
            ("Prix unitaire HT (FCFA)", donnees.get("prix_ht",0)),
            ("Sous-total HT (FCFA)", donnees.get("sous_total_ht",0)),
            (f"TVA ({donnees.get('tva_pct',0):.0f}%) (FCFA)", donnees.get("montant_tva",0)),
            ("Total TTC (FCFA)", donnees.get("total_ttc",0)),
        ]

        r_idx = 2
        for libelle, valeur in rows:
            if valeur == "" or valeur is None:
                continue
            ws.cell(row=r_idx, column=1, value=libelle).font = Font(bold=True)
            ws.cell(row=r_idx, column=2, value=valeur)
            for c in [1,2]:
                ws.cell(row=r_idx, column=c).alignment = Alignment(vertical="top", wrap_text=True)
            r_idx += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

# ── Traitement batch ─────────────────────────────────────────────────────────

def traiter_un_fichier(fichier_info: tuple) -> dict:
    idx, fichier, solde = fichier_info
    try:
        if not extension_autorisee(fichier.filename):
            return {"index": idx, "nom": fichier.filename,
                    "erreur": "Format non supporte. Seuls PNG et PDF sont acceptes."}
        img_bytes = fichier.read()
        fichier.seek(0)
        texte = extraire_texte(img_bytes, fichier.filename)
        if not texte.strip():
            return {"index": idx, "nom": fichier.filename, "erreur": "Aucun texte detecte"}
        donnees = parser_facture_complet(texte, fichier.filename, solde=solde)
        image_b64 = generer_image_facture(donnees)
        return {"index": idx, "nom": fichier.filename, "succes": True,
                "donnees": donnees, "image_b64": image_b64}
    except Exception as e:
        return {"index": idx, "nom": fichier.filename, "erreur": str(e)}

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html", tesseract_ok=TESSERACT_OK, couleurs=COULEURS_GESFI)

@app.route("/analyser_batch", methods=["POST"])
def analyser_batch():
    if not TESSERACT_OK:
        return jsonify({"erreur": "Tesseract non installe"}), 500
    if "fichiers" not in request.files:
        return jsonify({"erreur": "Aucun fichier recu"}), 400
    fichiers = request.files.getlist("fichiers")
    soldes = []
    for i in range(len(fichiers)):
        val = request.form.get(f"solde_{i}", "false").lower()
        soldes.append(val == "true")
    if len(fichiers) > 100:
        return jsonify({"erreur": "Maximum 100 fichiers autorises"}), 400
    if len(fichiers) == 0:
        return jsonify({"erreur": "Aucun fichier selectionne"}), 400

    resultats = []
    fichiers_a_traiter = [(i, f, soldes[i] if i < len(soldes) else False)
                          for i, f in enumerate(fichiers) if f.filename]
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(traiter_un_fichier, ft): ft for ft in fichiers_a_traiter}
        for future in as_completed(futures):
            resultats.append(future.result())
    resultats.sort(key=lambda x: x.get("index", 0))

    return jsonify({
        "total": len(resultats),
        "succes": len([r for r in resultats if r.get("succes")]),
        "erreurs": len([r for r in resultats if r.get("erreur")]),
        "resultats": resultats,
    })

@app.route("/telecharger_excel", methods=["POST"])
def telecharger_excel():
    data = request.get_json()
    if not data or "resultats" not in data:
        return jsonify({"erreur": "Donnees manquantes"}), 400
    excel_bytes = generer_excel_batch(data.get("resultats", []))
    return send_file(io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"factures_extraites_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

@app.route("/sante")
def sante():
    return jsonify({"status": "ok", "tesseract": TESSERACT_OK, "version": "4.0.0",
                    "formats_supportes": ["PNG", "PDF"], "timestamp": datetime.now().isoformat()})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
