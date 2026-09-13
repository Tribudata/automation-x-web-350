const $ = id => document.getElementById(id);
let records = [];
let jobId = null;

$("excelFile").addEventListener("change", async e => {
  const file = e.target.files[0];
  if (!file) return;
  $("fileName").textContent = file.name;
  $("processing").classList.remove("hidden");
  $("summary").classList.add("hidden");
  $("confirm").classList.add("hidden");
  setProgress(10, "Leyendo archivo…");

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {type:"array"});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
    setProgress(45, "Aplicando reglas de negocio…");

    records = rows.map((r, i) => normalize(r, i + 2));
    const valid = records.filter(r => r.errors.length === 0);
    const invalid = records.filter(r => r.errors.length > 0);

    setProgress(100, "Procesamiento terminado.");
    $("total").textContent = records.length;
    $("valid").textContent = valid.length;
    $("invalid").textContent = invalid.length;

    $("errors").innerHTML = invalid.length
      ? invalid.slice(0, 10).map(r => `<div>Fila ${r.row}: ${r.errors.join(" · ")}</div>`).join("")
        + (invalid.length > 10 ? `<div>… y ${invalid.length-10} errores más.</div>` : "")
      : "<div>✓ No se encontraron errores.</div>";

    $("summary").classList.remove("hidden");
    $("startBtn").disabled = valid.length === 0;
  } catch (err) {
    setProgress(0, "No fue posible procesar el archivo: " + err.message);
  }
});

function normalize(r, row) {
  // La configuración puede ampliarse por automatización.
  const nit = String(r.NIT ?? r.Identificacion ?? r.identificacion ?? "").trim();
  const periodo = String(r.AÑO ?? r.Año ?? r.Periodo ?? r.periodo ?? "").trim();
  const concepto = String(r.CONCEPTO ?? r.Concepto ?? r.campo ?? "").trim();
  const raw = r.VALOR ?? r.Valor ?? r.valor ?? "";
  const valor = Number(String(raw).replace(/[$.\s]/g, "").replace(/,/g, ""));
  const errors = [];
  if (!nit) errors.push("NIT/Identificación vacío");
  if (!periodo) errors.push("Período vacío");
  if (!concepto) errors.push("Concepto/campo vacío");
  if (!Number.isFinite(valor)) errors.push("Valor no numérico");
  return {row, nit, periodo, concepto, valor, errors};
}

$("startBtn").addEventListener("click", () => {
  $("confirm").classList.remove("hidden");
  $("confirm").scrollIntoView({behavior:"smooth"});
});

$("cancelBtn").addEventListener("click", () => $("confirm").classList.add("hidden"));

$("acceptBtn").addEventListener("click", () => {
  $("confirm").classList.add("hidden");
  $("status").classList.remove("hidden");
  jobId = "LOCAL-MVP-" + Date.now();
  $("jobId").textContent = "Trabajo preparado: " + jobId;
  setState("PREPARADO");
  log("✓ Validaciones completadas.");
  log("✓ Trabajo de automatización creado.");
  log("⏳ Pendiente de conectar el puente seguro con GitHub Actions.");
  log("ℹ️ En producción, aquí se solicitará la sesión de navegador remoto.");
  $("status").scrollIntoView({behavior:"smooth"});
});

function setProgress(p, text) {
  $("progressBar").style.width = p + "%";
  $("processingStatus").textContent = text;
}
function setState(text) { $("stateText").textContent = text; }
function log(text) { $("log").textContent += ($("log").textContent ? "\n" : "") + text; }
