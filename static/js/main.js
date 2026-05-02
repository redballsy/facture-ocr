const dropZone   = document.getElementById("dropZone");
const fileInput  = document.getElementById("fileInput");
const previewBox = document.getElementById("previewBox");
const previewImg = document.getElementById("previewImg");
const btnRemove  = document.getElementById("btnRemove");
const btnAnalyser = document.getElementById("btnAnalyser");
const btnText    = btnAnalyser.querySelector(".btn-text");
const btnLoader  = btnAnalyser.querySelector(".btn-loader");
const results    = document.getElementById("results");

let currentFile = null;
let currentData = null;

// ── Drag & Drop ───────────────────────────────────────────────────────────────

dropZone.addEventListener("dragover", e => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));

dropZone.addEventListener("drop", e => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) setFile(file);
});

dropZone.addEventListener("click", e => {
  if (e.target === dropZone || e.target.classList.contains("upload-icon") ||
      e.target.tagName === "P") {
    fileInput.click();
  }
});

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) setFile(fileInput.files[0]);
});

btnRemove.addEventListener("click", resetFile);

function setFile(file) {
  currentFile = file;
  const url = URL.createObjectURL(file);
  previewImg.src = url;
  dropZone.style.display = "none";
  previewBox.style.display = "block";
  btnAnalyser.disabled = false;
  results.style.display = "none";
  currentData = null;
}

function resetFile() {
  currentFile = null;
  currentData = null;
  previewImg.src = "";
  previewBox.style.display = "none";
  dropZone.style.display = "block";
  btnAnalyser.disabled = true;
  results.style.display = "none";
  fileInput.value = "";
}

// ── Analyser ──────────────────────────────────────────────────────────────────

btnAnalyser.addEventListener("click", async () => {
  if (!currentFile) return;

  // UI loading
  btnText.style.display = "none";
  btnLoader.style.display = "inline";
  btnAnalyser.disabled = true;

  const formData = new FormData();
  formData.append("fichier", currentFile);

  try {
    const res = await fetch("/analyser", { method: "POST", body: formData });
    const data = await res.json();

    if (!res.ok || data.erreur) {
      showToast(data.erreur || "Erreur serveur", "error");
      return;
    }

    currentData = data;
    currentData.nom_fichier = currentFile.name;
    afficherResultats(data);

  } catch (err) {
    showToast("Impossible de contacter le serveur.", "error");
  } finally {
    btnText.style.display = "inline";
    btnLoader.style.display = "none";
    btnAnalyser.disabled = false;
  }
});

function afficherResultats(d) {
  const f = d.fournisseur || {};
  const c = d.client || {};

  setText("r-numero",    d.numero_facture);
  setText("r-date-em",   d.date_emission);
  setText("r-date-ec",   d.date_echeance);
  setText("r-total",     d.total_ttc ? `${d.total_ttc} ${d.devise || ""}` : null);
  setText("r-four-nom",  f.nom);
  setText("r-four-adr",  f.adresse);
  setText("r-four-rccm", f.rccm);
  setText("r-four-cap",  f.capital);
  setText("r-four-tel",  f.telephone);
  setText("r-four-email",f.email);
  setText("r-cli-nom",   c.nom);
  setText("r-paiement",  d.mode_paiement);
  setText("r-sousTotal", d.sous_total_ht ? `${d.sous_total_ht} ${d.devise || ""}` : null);
  setText("r-tva",       d.tva_montant   ? `${d.tva_montant} ${d.devise || ""}`   : null);

  // Lignes
  const tbody = document.getElementById("linesBody");
  tbody.innerHTML = "";
  const lignes = d.lignes || [];
  if (lignes.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" class="no-lines">Aucune ligne détectée automatiquement — vérifier le texte OCR brut</td></tr>`;
  } else {
    lignes.forEach(l => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${esc(l.description)}</td>
        <td style="text-align:center">${esc(l.quantite)}</td>
        <td style="text-align:right">${esc(l.prix_unitaire_ht)}</td>
        <td style="text-align:center">${esc(l.tva_pct)}</td>
        <td style="text-align:right">${esc(l.total)}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  document.getElementById("rawText").textContent = d.texte_brut || "";
  document.getElementById("resultsMeta").textContent =
    `${d.mots_extraits || "?"} mots extraits · ${lignes.length} ligne(s) · ${currentFile.name}`;

  results.style.display = "block";
  results.scrollIntoView({ behavior: "smooth", block: "start" });
}

function setText(id, val) {
  document.getElementById(id).textContent = val || "—";
}

function esc(str) {
  if (!str) return "—";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ── Télécharger Excel ─────────────────────────────────────────────────────────

document.getElementById("btnDownload").addEventListener("click", async () => {
  if (!currentData) return;

  const btn = document.getElementById("btnDownload");
  btn.textContent = "⏳ Génération…";
  btn.disabled = true;

  try {
    const res = await fetch("/telecharger", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(currentData),
    });

    if (!res.ok) {
      showToast("Erreur lors de la génération Excel.", "error");
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const cd = res.headers.get("Content-Disposition") || "";
    const match = cd.match(/filename="?([^"]+)"?/);
    a.download = match ? match[1] : "facture.xlsx";
    a.href = url;
    a.click();
    URL.revokeObjectURL(url);
    showToast("Fichier Excel téléchargé ✓", "success");

  } catch (err) {
    showToast("Erreur réseau.", "error");
  } finally {
    btn.textContent = "↓ Télécharger le fichier Excel";
    btn.disabled = false;
  }
});

// ── Toast ─────────────────────────────────────────────────────────────────────

function showToast(msg, type = "info") {
  const existing = document.querySelector(".toast");
  if (existing) existing.remove();

  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  t.style.cssText = `
    position:fixed; bottom:24px; right:24px; z-index:9999;
    background:${type === "error" ? "#3a1a1a" : "#1a3a1a"};
    border:1px solid ${type === "error" ? "#ff6b6b" : "#c8ff57"};
    color:${type === "error" ? "#ff6b6b" : "#c8ff57"};
    padding:12px 20px; border-radius:10px;
    font-family:'DM Mono',monospace; font-size:13px;
    animation:fadeUp 0.3s ease;
    box-shadow:0 4px 20px rgba(0,0,0,0.4);
  `;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 4000);
}
