/* ========= Neutraliser "Script error." (bruit cross-origin de Word/Script Lab) ========= */
(function () {
  const isGeneric = (m)=> m==="Script error." || /(^|[^a-z])script error\.?/i.test(String(m||""));
  // capture-phase: bloque les handlers de Script Lab
  window.addEventListener("error", (e)=>{
    if(isGeneric(e?.message) && (!e?.filename || e.filename==="")){
      e.preventDefault(); e.stopImmediatePropagation(); return false;
    }
  }, true);
  // onerror fallback
  const prev = window.onerror;
  window.onerror = function(msg, src){
    if(isGeneric(msg) && (!src || src==="")) return true;
    return typeof prev==="function" ? prev.apply(this, arguments) : false;
  };
  // promesses
  window.addEventListener("unhandledrejection",(e)=>{
    const r=e?.reason, m=(r&& (r.message||r.toString?.()))||"";
    if(isGeneric(m)) e.preventDefault();
  });
})();

/* =================== Helpers statut =================== */
function imagesSetDone(done){
  const icon = document.getElementById("imagesState");
  const box  = document.getElementById("imagesStatus");
  if(!icon || !box) return;
  if(done){
    icon.textContent = "✔️";
    box.classList.remove("fail"); box.classList.add("ok");
    box.textContent = "✔️ Exercice réussi !";
  }else{
    icon.textContent = "❌";
    box.classList.remove("ok"); box.classList.add("fail");
    box.textContent = "❌ Exercice non encore réussi";
  }
  // Note: localStorage retiré car non supporté dans les artifacts Claude
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem("images_done", done ? "1" : "0");
    }
  } catch(e) {
    // Ignore silencieusement si localStorage n'est pas disponible
  }
}

// Restaurer état au chargement
(function(){ 
  try {
    if (typeof localStorage !== 'undefined') {
      imagesSetDone(localStorage.getItem("images_done")==="1"); 
    }
  } catch(e) {
    imagesSetDone(false);
  }
})();

/* =================== Gate =================== */
async function imagesStart() {
  let hasContent = false;

  // Étape 1 : vérifier si le document contient déjà du texte
  await Word.run(async (context) => {
    const body = context.document.body;
    const paras = body.paragraphs;
    paras.load("items/text");
    await context.sync();

    hasContent = paras.items.some(p => (p.text || "").trim().length > 0);
  });

  // Étape 3 : afficher l'exercice
  document.getElementById("images-gate").classList.add("is-hidden");
  document.getElementById("images-main").hidden = false;

  // Reset état
  imagesSetDone(false);
  const status = document.getElementById("imagesStatus");
  if (status) status.textContent = "❌ Exercice non encore réussi (document prêt)";
}

window.imagesStart = imagesStart;



/* =================== Validation (inline + flottant) =================== */
async function imagesValidate() {
  let okA = false, okB = false;
  const issues = [];

  try {
    await Word.run(async (ctx) => {
      const body  = ctx.document.body;
      const paras = body.paragraphs;
      const pics  = body.inlinePictures;

      // On charge ce qui est 100% supporté sur le web
      paras.load("items/text,items/alignment");
      pics.load("items");
      await ctx.sync();

      const hasInlineImage = pics.items.length > 0;
      if (!hasInlineImage) issues.push("aucune image en ligne trouvée");

      const hasRightText = paras.items.some(p =>
        ((p.alignment === "Right") || (p.alignment === 2)) &&
        (p.text || "").trim().length > 0
      );
      if (!hasRightText) issues.push("pas de paragraphe de texte aligné à droite");

      // Cas A : image inline + texte à droite
      okA = hasInlineImage && hasRightText;

      // Cas B (optionnel) : essayer shapes UNIQUEMENT si dispo
      let shapesSupported = false;
      try {
        // Petit test de présence + requirement set
        shapesSupported =
          Office.context.requirements.isSetSupported("WordApi", "1.3") &&
          typeof body.shapes !== "undefined";
      } catch (_) { shapesSupported = false; }

      if (shapesSupported) {
        try {
          const shapes = body.shapes;
          shapes.load("items/left,items.wrapType");
          await ctx.sync();

          const hasFloatingLeftShape = shapes.items.some(s => {
            const wrap = (s.wrapType || "").toString().toLowerCase();
            const left = Number(s.left);
            return wrap && wrap !== "inline" && !Number.isNaN(left) && left < 200;
          });

          const hasSomeText = paras.items.some(p => (p.text || "").trim().length > 0);
          okB = hasFloatingLeftShape && hasSomeText;
        } catch (_) {
          // On ignore si le web backend ne renvoie rien d’exploitable
          okB = false;
        }
      } else {
        // Sur Web : on ne bloque PAS l’utilisateur parce que shapes est absent
        okB = false;
      }

      const ok = okA || okB;
      imagesSetDone(ok);

      const status = document.getElementById("imagesStatus");
      if (ok) {
        status.textContent = okA
          ? "✔️ Exercice réussi (image en ligne + texte aligné à droite)."
          : "✔️ Exercice réussi (image flottante à gauche + texte).";
        status.className = "status ok";
      } else {
        status.textContent = "À corriger : " + [...new Set(issues)].join(" · ");
        status.className = "status fail";
      }
    });
  } catch (error) {
    console.error("Validation error (web-safe):", error);
    const status = document.getElementById("imagesStatus");
    status.textContent = "❌ Erreur pendant la validation (Word Online).";
    status.className = "status fail";
  }
}
window.imagesValidate = imagesValidate;
// Afficher/masquer le tutoriel
function imagesToggleTuto() {
  const bloc = document.getElementById('imagesTuto');
  const btn  = document.getElementById('imagesTutoBtn');
  if (!bloc || !btn) return;

  const isHidden = bloc.hasAttribute('hidden');
  if (isHidden) {
    bloc.removeAttribute('hidden');
    btn.textContent = "📘 Masquer le tutoriel";
  } else {
    bloc.setAttribute('hidden', '');
    btn.textContent = "📘 Afficher le tutoriel";
  }
}
window.imagesToggleTuto = imagesToggleTuto;