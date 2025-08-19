// Quand Office est prêt
Office.onReady(() => {
  document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("btnInsert").addEventListener("click", insertTwoColumnLayout);
    document.getElementById("btnAlt").addEventListener("click", insertAltParagraphRight);
  });
});

/** Petit helper pour clamp (%) */
function clampPct(v) {
  if (isNaN(v)) return 50;
  return Math.max(10, Math.min(90, v));
}

/** Convertit un fichier image en base64 */
// --- Utils robustes ---

function base64FromArrayBuffer(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  let binary = "";
  const chunk = 0x8000; // pour éviter les dépassements de pile
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/** Convertit un fichier image en base64 (sans prefix data:) */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    if (!file) return resolve(null);

    const allowed = ["image/png", "image/jpeg", "image/gif", "image/webp"];
    if (file.type && !allowed.includes(file.type)) {
      alert("Format non supporté. Choisis PNG/JPEG/GIF/WebP.");
      return resolve(null);
    }

    const reader = new FileReader();
    reader.onload = () => {
      try {
        // Méthode robuste: ArrayBuffer -> base64
        const buf = reader.result; // ArrayBuffer
        const base64 = base64FromArrayBuffer(buf);
        resolve(base64);
      } catch (e) {
        reject(e);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file); // ⚠️ on lit en ArrayBuffer (plus fiable)
  });
}

// Image PNG de test (petit carré 64x64) pour diagnostiquer si Word insère bien
const TEST_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAACXBIWXMAAAsSAAALEgHS3X78AAAAGXRFWHRTb2Z0d2FyZQBwYWludC5uZXQgNC4yLjGqq7UAAABhSURBVHic7cExAQAAAMKg9U9tCF8gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPgN5cQAAf3p4eQAAAABJRU5ErkJggg==";


/** Texte à gauche + image à droite (même ligne via tableau invisible) */
async function insertTwoColumnLayout() {
  const leftText = (document.getElementById("leftText").value || "").trim();
  const leftWidth = clampPct(parseInt(document.getElementById("leftWidth").value, 10) || 60);
  const rightWidth = clampPct(parseInt(document.getElementById("rightWidth").value, 10) || 40);
  const imgWidthPx = Math.max(40, parseInt(document.getElementById("imgWidth").value, 10) || 220);

  const file = document.getElementById("imageFile").files[0];
  const base64Img = await fileToBase64(file);

  await Word.run(async (context) => {
    const body = context.document.body;

    // Crée un tableau invisible
    const table = body.insertTable(1, 2, Word.InsertLocation.end, [[leftText || ""]]);
    ["InsideVertical","InsideHorizontal","Top","Bottom","Left","Right","Outside"]
      .forEach(b => table.getBorder(b).clear());

    try {
      table.columns.getItemAt(0).width = leftWidth;
      table.columns.getItemAt(1).width = rightWidth;
    } catch (e) { /* Word ajuste si non supporté */ }

    const row = table.rows.getFirst();
    const leftCell = row.cells.getItem(0).body;
    if (!leftText) {
      leftCell.clear();
      leftCell.insertParagraph("Texte à gauche…", Word.InsertLocation.start);
    }

    const rightCell = row.cells.getItem(1).body;
    if (base64Img) {
      const pic = rightCell.insertInlinePictureFromBase64(base64Img, Word.InsertLocation.start);
      pic.width = imgWidthPx;
    } else {
      rightCell.insertParagraph("(Aucune image sélectionnée)", Word.InsertLocation.start);
    }

    await context.sync();
  }).catch(err => alert("Erreur : " + err.message));
}

/** Alternative : texte puis image alignée à droite (en dessous) */
async function insertAltParagraphRight() {
  const leftText = (document.getElementById("leftText").value || "Texte à gauche").trim();
  const file = document.getElementById("imageFile").files[0];
  const base64Img = await fileToBase64(file);

  await Word.run(async (context) => {
    const body = context.document.body;

    const p = body.insertParagraph(leftText, Word.InsertLocation.end);
    p.alignment = "Left";

    if (base64Img) {
      const pic = body.insertInlinePictureFromBase64(base64Img, Word.InsertLocation.end);
      pic.parentParagraph.alignment = "Right";
    } else {
      body.insertParagraph("(Aucune image sélectionnée)", Word.InsertLocation.end).alignment = "Right";
    }

    await context.sync();
  }).catch(err => alert("Erreur : " + err.message));
}
