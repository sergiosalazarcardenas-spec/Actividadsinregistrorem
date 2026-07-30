/*
 * PANEL REM — configuración editable
 * ----------------------------------
 * Si cambia el encabezado de una planilla, agregue aquí otra alternativa.
 * La aplicación intenta primero encontrar el encabezado y, si no existe,
 * conserva la posición histórica: A, E, P, Q y AM.
 */
const CONFIG = {
  // La URL se define fuera del repositorio en config.local.js.
  // No ponga aquí una fuente con datos personales o clínicos.
  sourceUrl: String(globalThis.REM_SOURCE_URL ?? "").trim(),
  columns: {
    id: {
      label: "ID de atención",
      aliases: ["id", "aten id", "id atencion", "id atención", "identificador", "folio", "rut", "run"],
      fallbackColumn: "A",
      required: true,
    },
    time: {
      label: "Hora de atención",
      aliases: ["hora atend", "hora atencion", "hora atención", "hora", "hora de atencion", "hora de atención"],
      fallbackColumn: "E",
    },
    staff: {
      label: "Funcionario",
      aliases: ["funcionario", "profesional", "nombre funcionario", "nombre profesional", "prestador"],
      fallbackColumn: "P",
    },
    instrument: {
      label: "Instrumento",
      aliases: ["instrumento", "tipo instrumento", "nombre instrumento"],
      fallbackColumn: "Q",
    },
    activity: {
      label: "Actividad o procedimiento",
      aliases: ["actividad y o procedimiento", "actividad y/o procedimiento", "actividad", "procedimiento", "prestacion", "prestación"],
      fallbackColumn: "R",
    },
    rem: {
      label: "Concepto / detalle REM",
      aliases: ["concepto", "detalle rem", "estado rem", "rem", "descripcion rem", "descripción rem"],
      fallbackColumn: "AM",
      required: true,
    },
  },
  // Esta lista conserva la regla del HTML original. Agregue frases aquí si la codificación cambia.
  notCountedRemPhrases: ["actividad no contabilizada en rem", "actividad no contabilizada en ram"],
  previewLimit: 150,
  chartLimit: 8,
};

let rawData = [];
let filteredData = [];
let originalHeaders = [];
let originalRows = [];
let detectedColumns = {};
let mainChart = null;
let chartBy = "instrument";

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

document.addEventListener("DOMContentLoaded", () => {
  $("#loadButton").addEventListener("click", loadGoogleSheet);
  ["#filterDuplicate", "#filterStaff"].forEach((selector) => {
    $(selector).addEventListener("change", applyFilters);
  });
  $("#filterInstrument").addEventListener("change", onInstrumentChange);
  $("#filterSearch").addEventListener("input", debounce(applyFilters, 180));
  $("#clearFilters").addEventListener("click", clearFilters);
  $("#exportButton").addEventListener("click", exportFilteredData);
  $$("[data-chart-by]").forEach((button) => {
    button.addEventListener("click", () => {
      setChartBy(button.dataset.chartBy);
      renderChart(filteredData);
    });
  });
});

async function loadGoogleSheet() {
  const rawUrl = CONFIG.sourceUrl;
  let source;
  try {
    if (!rawUrl) throw new Error("Falta configurar la fuente en config.local.js.");
    source = buildGoogleSheetSource(rawUrl);
  } catch (error) {
    showToast(error.message, true);
    return;
  }

  setLoading(true);
  setSourceStatus("Leyendo Google Sheets…", "loading");
  try {
    if (typeof XLSX === "undefined") throw new Error("No se pudo cargar la biblioteca de lectura. Revise su conexión a internet.");

    const response = await fetch(source.dataUrl, { cache: "no-store" });
    if (!response.ok) throw new Error(`Google Sheets respondió con código ${response.status}.`);
    let rows;
    if (source.format === "tsv") {
      rows = parseDelimitedText(await response.text(), "\t");
    } else {
      const fileContent = await response.arrayBuffer();
      const workbook = XLSX.read(fileContent, { type: "array", raw: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    }

    if (rows.length < 2) throw new Error("La pestaña no contiene filas de datos después de los encabezados.");
    hydrateData(rows);
    updateSelectors();
    updateKPIs();
    applyFilters();
    setSourceStatus(`Datos cargados · ${source.label}`, "live");
    showToast(`${rawData.length.toLocaleString("es-CL")} registros actualizados.`);
  } catch (error) {
    console.error(error);
    setSourceStatus("No fue posible leer la fuente", "error");
    showToast(`${error.message} Verifique que la pestaña esté compartida como lector con enlace.`, true);
  } finally {
    setLoading(false);
  }
}

function buildGoogleSheetSource(value) {
  if (!value) throw new Error("Pegue el enlace de una hoja de Google Sheets.");
  let parsedUrl;
  try {
    parsedUrl = new URL(value);
  } catch {
    throw new Error("El enlace de Google Sheets no es válido.");
  }

  // Enlace de una hoja publicada: .../spreadsheets/d/e/PUBLICATION_ID/pub?output=tsv
  // Se conserva la misma pestaña publicada y se fuerza salida TSV para leerla con seguridad.
  const publishedMatch = parsedUrl.pathname.match(/^\/spreadsheets\/d\/e\/[^/]+\/pub$/);
  if (parsedUrl.hostname === "docs.google.com" && publishedMatch) {
    parsedUrl.searchParams.set("output", "tsv");
    return { dataUrl: parsedUrl.toString(), format: "tsv", label: "hoja publicada (TSV)" };
  }

  const idMatch = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  const directId = value.match(/^[a-zA-Z0-9_-]{25,}$/);
  const sheetId = idMatch?.[1] || directId?.[0];
  if (!sheetId) throw new Error("No encontré el identificador de la planilla en ese enlace.");

  let gid = "0";
  const gidMatch = value.match(/[?#&]gid=(\d+)/);
  if (gidMatch) gid = gidMatch[1];
  return {
    dataUrl: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${encodeURIComponent(gid)}`,
    format: "csv",
    label: `pestaña gid=${gid}`,
  };
}

// Lee CSV o TSV respetando comillas, tabulaciones y saltos de línea dentro de una celda.
function parseDelimitedText(text, delimiter) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (quoted) {
      if (char === '"' && text[index + 1] === '"') {
        cell += '"';
        index += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        cell += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === delimiter) {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (char !== "\r") {
      cell += char;
    }
  }
  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

function hydrateData(rows) {
  originalHeaders = rows[0].map((cell, index) => String(cell || `Columna ${index + 1}`).trim());
  originalRows = rows.slice(1);
  detectedColumns = detectColumns(originalHeaders);

  const missing = Object.entries(CONFIG.columns)
    .filter(([key, config]) => config.required && detectedColumns[key] === -1)
    .map(([, config]) => config.label);
  if (missing.length) throw new Error(`No se detectaron las columnas obligatorias: ${missing.join(", ")}.`);

  const idStats = new Map();
  const activityStatsById = new Map();
  originalRows.forEach((row) => {
    const id = getCell(row, "id").trim();
    const remText = getCell(row, "rem");
    if (!id) return;

    const item = idStats.get(id) || { total: 0, notCounted: 0 };
    item.total += 1;
    if (isNotCountedRem(remText)) item.notCounted += 1;
    idStats.set(id, item);

    // Una misma atención puede tener actividades diferentes. Para detectar una
    // repetición real, se agrupa además por la actividad/procedimiento y no
    // solo por el ID. Si no existe esa columna se usa el detalle REM como
    // respaldo para no interrumpir el análisis.
    const activityKey = getActivityKey(row, remText);
    const activitiesForId = activityStatsById.get(id) || new Map();
    const activityItem = activitiesForId.get(activityKey) || { total: 0, notCounted: 0 };
    activityItem.total += 1;
    if (isNotCountedRem(remText)) activityItem.notCounted += 1;
    activitiesForId.set(activityKey, activityItem);
    activityStatsById.set(id, activitiesForId);
  });

  rawData = originalRows.map((row, index) => {
    const id = getCell(row, "id").trim();
    const remDetail = getCell(row, "rem").trim();
    const group = idStats.get(id);
    const activity = getCell(row, "activity").trim();
    const activityGroup = activityStatsById.get(id)?.get(getActivityKey(row, remDetail));
    const isDuplicate = Boolean(id && group?.total > 1);
    const isAllNotCountedId = Boolean(isDuplicate && group.total === group.notCounted);
    const isRepeatedUnregisteredActivity = Boolean(
      id && activityGroup?.total > 1 && activityGroup.total === activityGroup.notCounted,
    );
    // Mantiene el criterio histórico (ID completo sin REM) y suma los casos
    // donde solo una actividad se repite dos o más veces sin REM, aunque ese
    // mismo ID incluya otras actividades registradas en REM.
    const isStrictNotCountedDup = isAllNotCountedId || isRepeatedUnregisteredActivity;
    const notCounted = isNotCountedRem(remDetail);

    return {
      originalRowIndex: index,
      rowNumber: index + 2,
      id,
      time: formatTime(getCell(row, "time")),
      instrument: getCell(row, "instrument").trim() || "No especificado",
      staff: getCell(row, "staff").trim() || "No especificado",
      activity: activity || remDetail,
      remDetail,
      isDuplicate,
      isStrictNotCountedDup,
      isRepeatedUnregisteredActivity,
      isNotCounted: notCounted,
      remClassification: notCounted ? "No contabilizada REM" : "Contabilizada / otro concepto",
    };
  });

  renderColumnMapping();
  renderDataQuality();
}

function detectColumns(headers) {
  const result = {};
  for (const [key, config] of Object.entries(CONFIG.columns)) {
    const aliases = config.aliases.map(normalizeText);
    let index = headers.findIndex((header) => aliases.includes(normalizeText(header)));
    if (index === -1) {
      index = headers.findIndex((header) => {
        const normalizedHeader = normalizeText(header);
        return aliases.some((alias) => normalizedHeader.includes(alias));
      });
    }
    if (index === -1) {
      const fallbackIndex = columnLetterToIndex(config.fallbackColumn);
      index = headers.length > fallbackIndex ? fallbackIndex : -1;
    }
    result[key] = index;
  }
  return result;
}

function getCell(row, key) {
  const index = detectedColumns[key];
  return index >= 0 && row[index] !== undefined && row[index] !== null ? String(row[index]) : "";
}

function getActivityKey(row, remText) {
  const activity = getCell(row, "activity").trim();
  return normalizeText(activity) || normalizeText(remText);
}

function isNotCountedRem(value) {
  const normalized = normalizeText(value);
  return CONFIG.notCountedRemPhrases.some((phrase) => normalized.includes(normalizeText(phrase)));
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function columnLetterToIndex(letters) {
  return letters.split("").reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function formatTime(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const numeric = Number(raw.replace(",", "."));
  if (!Number.isNaN(numeric) && numeric >= 0 && numeric < 1) {
    const seconds = Math.round(numeric * 86400);
    const hours = Math.floor(seconds / 3600) % 24;
    const minutes = Math.floor((seconds % 3600) / 60);
    const secs = seconds % 60;
    return [hours, minutes, secs].map((part) => String(part).padStart(2, "0")).join(":");
  }
  return raw;
}

function updateSelectors() {
  updateSelect("#filterInstrument", rawData.map((item) => item.instrument), "Todos los instrumentos");
  updateStaffForSelectedInstrument();
}

function updateSelect(selector, values, allLabel) {
  const select = $(selector);
  const previous = select.value;
  const uniqueValues = [...new Set(values)].sort((a, b) => a.localeCompare(b, "es"));
  select.replaceChildren(new Option(allLabel, "ALL"), ...uniqueValues.map((value) => new Option(value, value)));
  select.value = uniqueValues.includes(previous) ? previous : "ALL";
}

// Cascada Instrumento → Funcionario: evita ofrecer funcionarios sin registros
// en el instrumento que acaba de escoger el usuario.
function updateStaffForSelectedInstrument() {
  const selectedInstrument = $("#filterInstrument").value;
  const staffPool = selectedInstrument === "ALL"
    ? rawData
    : rawData.filter((item) => item.instrument === selectedInstrument);
  updateSelect("#filterStaff", staffPool.map((item) => item.staff), "Todos los funcionarios");
}

function onInstrumentChange() {
  updateStaffForSelectedInstrument();
  // Un único instrumento ya no se describe mejor con una barra propia: se
  // despliega automáticamente por los funcionarios que lo componen.
  setChartBy($("#filterInstrument").value === "ALL" ? "instrument" : "staff");
  applyFilters();
}

function setChartBy(value) {
  chartBy = value;
  $$("[data-chart-by]").forEach((button) => button.classList.toggle("is-active", button.dataset.chartBy === value));
}

function applyFilters() {
  const filters = {
    duplicate: $("#filterDuplicate").value,
    instrument: $("#filterInstrument").value,
    staff: $("#filterStaff").value,
    search: normalizeText($("#filterSearch").value),
  };
  filteredData = rawData.filter((item) => {
    if (filters.duplicate === "DUPLICATED" && !item.isDuplicate) return false;
    if (filters.duplicate === "DUPLICATED_STRICT_REM" && !item.isStrictNotCountedDup) return false;
    if (filters.instrument !== "ALL" && item.instrument !== filters.instrument) return false;
    if (filters.staff !== "ALL" && item.staff !== filters.staff) return false;
    if (filters.search && !normalizeText(item.id).includes(filters.search)) return false;
    return true;
  });
  renderTable(filteredData);
  renderCrossTable(filteredData);
  renderChart(filteredData);
  $("#exportButton").disabled = filteredData.length === 0;
  $("#selectionSummary").textContent = rawData.length
    ? `${filteredData.length.toLocaleString("es-CL")} de ${rawData.length.toLocaleString("es-CL")} registros cumplen los filtros actuales.`
    : "Sin selección activa";
}

function clearFilters() {
  $("#filterDuplicate").value = "ALL";
  $("#filterInstrument").value = "ALL";
  $("#filterStaff").value = "ALL";
  $("#filterSearch").value = "";
  updateStaffForSelectedInstrument();
  setChartBy("instrument");
  applyFilters();
}

function updateKPIs() {
  const ids = new Set(rawData.filter((item) => item.id).map((item) => item.id));
  // Repeticiones adicionales: cada ID aporta una primera fila única y las
  // siguientes se contabilizan como registros múltiples.
  const multipleRecordRows = Math.max(rawData.length - ids.size, 0);
  const strict = rawData.filter((item) => item.isStrictNotCountedDup).length;
  $("#metricTotalRows").textContent = rawData.length.toLocaleString("es-CL");
  $("#metricTotalRowsNote").textContent = "Filas de datos válidas";
  $("#metricUnique").textContent = ids.size.toLocaleString("es-CL");
  $("#metricDuplicates").textContent = multipleRecordRows.toLocaleString("es-CL");
  $("#metricStrictRem").textContent = strict.toLocaleString("es-CL");
}

function renderTable(data) {
  const body = $("#resultsTableBody");
  const shown = Math.min(data.length, CONFIG.previewLimit);
  $("#tableStatus").textContent = data.length
    ? `Mostrando ${shown.toLocaleString("es-CL")} de ${data.length.toLocaleString("es-CL")}`
    : "Sin coincidencias";
  if (!data.length) {
    body.innerHTML = '<tr><td colspan="7" class="empty-cell">No hay registros que cumplan los filtros seleccionados.</td></tr>';
    return;
  }
  body.innerHTML = data.slice(0, CONFIG.previewLimit).map((item) => {
    const duplicatePill = item.isStrictNotCountedDup
      ? '<span class="pill pill--strict">Repetición 100% sin REM</span>'
      : item.isDuplicate
        ? '<span class="pill pill--duplicate">Duplicado</span>'
        : '<span class="pill pill--unique">Único</span>';
    return `<tr>
      <td>#${item.rowNumber}</td>
      <td class="id-cell">${valueOrEmpty(item.id)}</td>
      <td>${valueOrEmpty(item.time)}</td>
      <td>${duplicatePill}</td>
      <td>${escapeHtml(item.instrument)}</td>
      <td>${escapeHtml(item.staff)}</td>
      <td class="concept-cell ${item.isNotCounted ? "rem-mark" : ""}" title="${escapeHtml(item.remDetail)}">${valueOrEmpty(item.remDetail)}</td>
    </tr>`;
  }).join("");
}

function renderCrossTable(data) {
  const grouped = new Map();
  data.forEach((item) => {
    const value = grouped.get(item.instrument) || { notCounted: 0, other: 0 };
    if (item.isNotCounted) value.notCounted += 1;
    else value.other += 1;
    grouped.set(item.instrument, value);
  });
  const rows = [...grouped.entries()]
    .map(([instrument, values]) => ({ instrument, ...values, total: values.notCounted + values.other }))
    .sort((a, b) => b.total - a.total || a.instrument.localeCompare(b.instrument, "es"));
  $("#crossTotal").textContent = `${data.length.toLocaleString("es-CL")} registros`;
  $("#crossTableBody").innerHTML = rows.length ? rows.map((row) => {
    const rate = row.total ? (row.notCounted / row.total) * 100 : 0;
    return `<tr><td>${escapeHtml(row.instrument)}</td><td>${row.notCounted.toLocaleString("es-CL")}</td><td>${row.other.toLocaleString("es-CL")}</td><td><strong>${row.total.toLocaleString("es-CL")}</strong></td><td class="percentage">${formatPercent(rate)}</td></tr>`;
  }).join("") : '<tr><td colspan="5" class="empty-cell">Sin datos para cruzar.</td></tr>';
}

function renderChart(data) {
  const empty = $("#chartEmpty");
  const chartTitle = $("#chart-title");
  if (mainChart) mainChart.destroy();
  if (!data.length || typeof Chart === "undefined") {
    empty.style.display = "grid";
    return;
  }
  empty.style.display = "none";
  const selectedInstrument = $("#filterInstrument").value;
  const key = chartBy === "instrument" ? "instrument" : chartBy === "staff" ? "staff" : "remClassification";
  const title = chartBy === "instrument"
    ? "Registros por instrumento"
    : chartBy === "staff" && selectedInstrument !== "ALL"
      ? `Funcionarios del instrumento: ${selectedInstrument}`
      : chartBy === "staff"
        ? "Registros por funcionario"
        : "Clasificación REM";
  chartTitle.textContent = title;
  const counts = new Map();
  data.forEach((item) => counts.set(item[key], (counts.get(item[key]) || 0) + 1));
  const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, CONFIG.chartLimit);
  const palette = ["#087d78", "#3976a8", "#d8783c", "#795fc2", "#7c9a3d", "#c75e77", "#19758d", "#b98527"];
  mainChart = new Chart($("#mainChart"), {
    type: "bar",
    data: { labels: sorted.map(([label]) => label), datasets: [{ data: sorted.map(([, value]) => value), backgroundColor: palette, borderRadius: 7, borderSkipped: false }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { backgroundColor: "#0c2c3a", padding: 11, displayColors: false } },
      scales: { x: { grid: { display: false }, ticks: { color: "#64736d", maxRotation: 0, autoSkip: false, callback: function (_, index) { const label = this.getLabelForValue(index); return label.length > 18 ? `${label.slice(0, 18)}…` : label; } } }, y: { beginAtZero: true, ticks: { precision: 0, color: "#64736d" }, grid: { color: "#e6ede8" } } },
    },
  });
}

function renderColumnMapping() {
  const labels = Object.entries(CONFIG.columns).map(([key, config]) => {
    const index = detectedColumns[key];
    return index === -1 ? `${config.label}: no detectada` : `${config.label}: ${originalHeaders[index]}`;
  });
  $("#columnMapping").textContent = labels.join(" · ");
}

function renderDataQuality() {
  const blankIds = rawData.filter((item) => !item.id).length;
  const fallbackColumns = Object.entries(CONFIG.columns)
    .filter(([key, config]) => detectedColumns[key] === columnLetterToIndex(config.fallbackColumn) && !config.aliases.map(normalizeText).includes(normalizeText(originalHeaders[detectedColumns[key]])))
    .map(([, config]) => config.label);
  const notices = [];
  if (blankIds) notices.push(`${blankIds.toLocaleString("es-CL")} fila(s) no tienen ID; no participan en la detección de duplicados.`);
  if (fallbackColumns.length) notices.push(`Se usó la posición histórica para: ${fallbackColumns.join(", ")}. Conviene revisar los encabezados.`);
  const panel = $("#dataQuality");
  panel.hidden = !notices.length;
  panel.textContent = notices.join(" ");
}

function exportFilteredData() {
  if (!filteredData.length || typeof XLSX === "undefined") return;
  const addedHeaders = ["Clasificación REM", "Estado de duplicidad"];
  const toExport = (item) => [
    ...originalRows[item.originalRowIndex],
    item.remClassification,
    item.isStrictNotCountedDup ? "Repetición 100% sin REM" : item.isDuplicate ? "Duplicado" : "Único",
  ];
  const workbook = XLSX.utils.book_new();
  const headers = [...originalHeaders, ...addedHeaders];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...filteredData.map(toExport)]), "Consolidado filtrado");

  const byStaff = new Map();
  filteredData.forEach((item) => {
    const group = byStaff.get(item.staff) || [];
    group.push(toExport(item));
    byStaff.set(item.staff, group);
  });
  [...byStaff.entries()].forEach(([staff, rows], index) => {
    const baseName = sanitizeSheetName(staff) || "Profesional";
    const name = uniqueSheetName(workbook, baseName, index);
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...rows]), name);
  });
  XLSX.writeFile(workbook, "Reporte_REM_filtrado.xlsx");
}

function sanitizeSheetName(value) { return String(value).replace(/[:\\/?*\[\]]/g, "").trim().slice(0, 31); }
function uniqueSheetName(workbook, proposed, index) {
  let name = proposed || `Profesional ${index + 1}`;
  let count = 1;
  while (workbook.SheetNames.includes(name)) {
    const suffix = `_${count++}`;
    name = `${proposed.slice(0, 31 - suffix.length)}${suffix}`;
  }
  return name;
}

function valueOrEmpty(value) { return value ? escapeHtml(value) : '<span class="muted">Sin dato</span>'; }
function escapeHtml(value) { return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;"); }
function formatPercent(value) { return `${value.toLocaleString("es-CL", { maximumFractionDigits: 1 })}%`; }
function debounce(callback, delay) { let id; return (...args) => { clearTimeout(id); id = setTimeout(() => callback(...args), delay); }; }

function setLoading(isLoading) {
  const button = $("#loadButton");
  button.disabled = isLoading;
  button.classList.toggle("is-loading", isLoading);
  button.querySelector(".button__label").textContent = isLoading ? "Actualizando" : "Actualizar datos";
}
function setSourceStatus(message, mode) {
  $("#sourceStatus").textContent = message;
  $("#sourceDot").className = `status-dot${mode === "live" ? " is-live" : mode === "loading" ? " is-loading" : ""}`;
}
let toastTimer;
function showToast(message, isError = false) {
  const toast = $("#toast");
  toast.textContent = message;
  toast.classList.toggle("is-error", isError);
  toast.classList.add("is-visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 5600);
}
