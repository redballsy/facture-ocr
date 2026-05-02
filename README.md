# FactureOCR 📊

**Application web locale pour extraire les données de factures (images) vers Excel.**

100% local · Sans API · Sans abonnement · Tesseract OCR

---

## Fonctionnalités

- Upload d'image de facture (JPG, PNG, BMP, TIFF, WEBP)
- Extraction OCR via Tesseract (moteur local)
- Détection automatique des champs : numéro, dates, fournisseur, client, lignes, totaux
- Téléchargement Excel (.xlsx) structuré et mis en forme
- Onglet "Texte brut OCR" pour vérification

## Installation

### 1. Cloner le dépôt

```bash
git clone https://github.com/TON_USERNAME/facture-ocr.git
cd facture-ocr
```

### 2. Installer Tesseract OCR

**Windows :**
- Télécharger sur https://github.com/UB-Mannheim/tesseract/wiki
- Pendant l'installation, cocher **"French language data"**
- Laisser le chemin par défaut : `C:\Program Files\Tesseract-OCR\`

**Linux :**
```bash
sudo apt install tesseract-ocr tesseract-ocr-fra
```

**macOS :**
```bash
brew install tesseract tesseract-lang
```

### 3. Installer les dépendances Python

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/macOS
source venv/bin/activate

pip install -r requirements.txt
```

### 4. Lancer l'application

```bash
python app.py
```

Ouvre ensuite : http://localhost:5000

## Structure du projet

```
facture-ocr/
├── app.py              # Serveur Flask + logique OCR + génération Excel
├── requirements.txt    # Dépendances Python
├── templates/
│   └── index.html      # Interface web
├── static/
│   ├── css/style.css   # Styles
│   └── js/main.js      # Logique frontend
└── README.md
```

## Technologies

- **Flask** — serveur web Python
- **Tesseract OCR** — moteur de reconnaissance de texte
- **OpenCV** — prétraitement d'image
- **OpenPyXL** — génération de fichiers Excel

## Licence

MIT
