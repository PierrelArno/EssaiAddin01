// ===============================
//  Word Add-in — Texte & Image
//  (Full API Word — aucun tableau)
// ===============================

/* ==== Silencer "Script error." (Word/Office cross-origin noise) ==== */
(function () {
  const isGenericMsg = (msg) =>
    msg === "Script error." || (typeof msg === "string" && /(^|[^a-z])script error\.?/i.test(msg));

  // Capture-phase listener: stoppe les handlers en aval (ex: handleError)
  window.addEventListener(
    "error",
    (e) => {
      const msg = e?.message || "";
      const src = e?.filename || "";
      if (isGenericMsg(msg) && (!src || src === "")) {
        e.preventDefault();
        e.stopImmediatePropagation();
        return false;
      }
    },
    true
  );

  // window.onerror shield
  const prevOnError = window.onerror;
  window.onerror = function (msg, src) {
    if (isGenericMsg(msg) && (!src || src === "")) return true;
    if (typeof prevOnError === "function") return prevOnError.apply(this, arguments);
    return false;
  };

  // Unhandled rejections (génériques)
  window.addEventListener("unhandledrejection", (e) => {
    const r = e?.reason;
    const msg = (r && (r.message || r.toString?.())) || "";
    if (isGenericMsg(msg)) e.preventDefault();
  });

  // Helper d'exécution sûre pour les boutons
  window.__safeRun = async (fn) => {
    try { await fn(); }
    catch (err) {
      console.error("[SoftError]", err && err.stack ? err.stack : err);
      alert("Une erreur est survenue. Détails en console.");
    }
  };
})();

/** Utils */
function clampNum(v, min, max, fallback) {
  const n = Number(v);
  if (Number.isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

/** Convertit un fichier image en base64 (sans préfixe data:) + mime */
function fileToBase64AndMime(file) {
  return new Promise((resolve) => {
    if (!file) return resolve({ base64: null, mime: "image/png" });

    const allowed = ["image/png", "image/jpeg", "image/gif"];
    const mime = allowed.includes(file.type) ? file.type : "image/png";
    if (file.type && !allowed.includes(file.type)) {
      alert("Format non supporté. Utilise PNG / JPEG / GIF.");
      return resolve({ base64: null, mime });
    }

    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || "");
      const i = dataUrl.indexOf(",");
      const base64 = i === -1 ? null : dataUrl.slice(i + 1);
      resolve({ base64, mime });
    };
    reader.onerror = () => resolve({ base64: null, mime });
    reader.readAsDataURL(file);
  });
}

// PNG 64×64 de test (si aucune image n'est fournie)
const TEST_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAACXBIWXMAAAsSAAALEgHS3X78AAAAGXRFWHRTb2Z0d2FyZQBwYWludC5uZXQgNC4yLjGqq7UAAABhSURBVHic7cExAQAAAMKg9U9tCF8gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPgN5cQAAf3p4eQAAAABJRU5ErkJggg==";

/** Nettoie un texte pour un paragraphe Word */
function sanitizeText(s) {
  return (s || "").replace(/\r\n|\r|\n/g, " ");
}

/** Insert: paragraphe texte (gauche), puis paragraphe image (droite). AUCUN tableau. */
async function insertTextAndImage() {
  const leftTextRaw = (document.getElementById("leftText").value || "").trim();
  const leftText = sanitizeText(leftTextRaw || "Texte à gauche…");
  const imgWidthPx = clampNum(document.getElementById("imgWidth").value, 40, 4096, 220);

  const file = document.getElementById("imageFile").files[0];
  const { base64 } = await fileToBase64AndMime(file);
  const imgBase64 = base64 || TEST_PNG_BASE64;

  await Word.run(async (context) => {
    const body = context.document.body;

    // 1) Paragraphe pour le texte (gauche)
    const pText = body.insertParagraph(leftText, Word.InsertLocation.end);
    pText.alignment = "Left";

    // 2) Paragraphe pour l'image (droite)
    const pImg = body.insertParagraph("", Word.InsertLocation.end);
    const pic = pImg.insertInlinePictureFromBase64(imgBase64, Word.InsertLocation.end);
    pic.width = imgWidthPx;
    pImg.alignment = "Right";

    await context.sync();
  });
}

/** Vider le document (utilitaire) */
async function clearDocument() {
  await Word.run(async (context) => {
    context.document.body.clear();  // Word garde juste le paragraphe final (vide, obligatoire)
    await context.sync();
  });
}


// Expose pour les onclick HTML
window.insertTextAndImage = insertTextAndImage;
window.clearDocument = clearDocument;

function startProject(){
  // masque l’intro
  document.getElementById('intro-screen').classList.add('is-hidden');

  // affiche l’UI du projet
  const main = document.getElementById('project-main');
  main.hidden = false;

  // focus sur le premier champ utile
  const first = document.getElementById('leftText') || main.querySelector('textarea, input, button');
  if (first) first.focus();
}
window.startProject = startProject;