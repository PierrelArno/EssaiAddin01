Office.onReady(() => {
  // Puis attendre que le DOM soit prêt
  document.addEventListener("DOMContentLoaded", () => {
    const btnInsert = document.getElementById("btnInsert");
    const btnAlt = document.getElementById("btnAlt");

    if (btnInsert) btnInsert.addEventListener("click", insertTwoColumnLayout);
    if (btnAlt) btnAlt.addEventListener("click", insertAltParagraphRight);
  });
});

/** Convertit un File (image) en base64 (sans le préfixe data:) */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    if (!file) return resolve(null);
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result; // "data:image/png;base64,XXXX"
      if (typeof result === "string") {
        const base64 = result.split(",")[1]; // garder seulement la partie base64
        resolve(base64);
      } else {
        resolve(null);
      }
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/** Insère un tableau 1x2 invisible : texte (gauche) + image (droite). */
async function insertTwoColumnLayout() {
  const leftText = (document.getElementById("leftText").value || "").trim();
  const leftWidth = clampPct(parseInt(document.getElementById("leftWidth")?.value, 10) || 60);
  const rightWidth = clampPct(parseInt(document.getElementById("rightWidth")?.value, 10) || 40);
  const imgWidthPx = Math.max(40, parseInt(document.getElementById("imgWidth")?.value, 10) || 220);

  const fileInput = document.getElementById("imageFile");
  const file = fileInput && fileInput.files && fileInput.files[0];
  const base64Img = await fileToBase64(file);

  await Word.run(async (context) => {
    const body = context.document.body;

    // Crée un tableau 1 ligne / 2 colonnes
    const table = body.insertTable(1, 2, Word.InsertLocation.end, [[leftText || ""]]);

    // Enlève toutes les bordures (tableau invisible)
    table.getBorder("InsideVertical").clear();
    table.getBorder("InsideHorizontal").clear();
    table.getBorder("Top").clear();
    table.getBorder("Bottom").clear();
    table.getBorder("Left").clear();
    table.getBorder("Right").clear();
    table.getBorder("Outside").clear();

    // Largeur des colonnes (selon build, Word peut ignorer les %)
    try {
      table.columns.getItemAt(0).width = leftWidth;   // %
      table.columns.getItemAt(1).width = rightWidth;  // %
    } catch (e) {
      // Si non supporté, Word fera l'ajustement auto.
    }

    // Cellule gauche : texte
    const row = table.rows.getFirst();
    const leftCellBody = row.cells.getItemAt(0).body;
    if (!leftText) {
      leftCellBody.clear();
      leftCellBody.insertParagraph("Texte à gauche…", Word.InsertLocation.start);
    }

    // Cellule droite : image ou placeholder
    const rightCellBody = row.cells.getItemAt(1).body;
    if (base64Img) {
      const pic = rightCellBody.insertInlinePictureFromBase64(base64Img, Word.InsertLocation.start);
      pic.width = imgWidthPx;
    } else {
      rightCellBody.insertParagraph("(Aucune image sélectionnée)", Word.InsertLocation.start);
    }

    await context.sync();
  }).catch((err) => {
    console.error(err);
    alert("Erreur lors de l'insertion : " + err.message);
  });
}