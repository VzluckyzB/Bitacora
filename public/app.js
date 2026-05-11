/* =========================================================
   KIDDE — BITÁCORA DE ACCESO
   app.js — Pure Vanilla JS + SheetJS
   ========================================================= */

"use strict";

/* ── State ─────────────────────────────────────────────── */
const state = {
  workbook:    null,   // XLSX workbook loaded from file
  fileHandle:  null,   // FileSystemFileHandle — para sobreescribir el mismo archivo
  fileName:    null,   // Nombre original del archivo cargado
  records:     [],     // pending rows [ {empresa,fechaIngreso,…} ]
  searchQuery: "",
};

// ¿El navegador soporta File System Access API?
const FS_API = "showOpenFilePicker" in window;

const HEADERS = [
  "Empresa",
  "Fecha Ingreso",
  "Hora Ingreso",
  "Matrícula",
  "Tipo Licencia",
  "Fecha Salida",
  "Hora Salida",
];

/* ── DOM refs ───────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);
const fileInput    = $("fileInput");
const uploadZone   = $("uploadZone");
const fileStatus   = $("fileStatus");
const accessForm   = $("accessForm");
const clearBtn     = $("clearBtn");
const saveBtn      = $("saveBtn");
const formMessage  = $("formMessage");
const searchInput  = $("searchInput");
const exportBtn    = $("exportBtn");
const recordsBody  = $("recordsBody");
const recordCount  = $("recordCount");
const themeToggle  = $("themeToggle");

/* ── Theme ──────────────────────────────────────────────── */
(function initTheme() {
  const saved = localStorage.getItem("kidde-theme") || "light-theme";
  document.body.className = saved;
  themeToggle.textContent = saved === "dark-theme" ? "🌙" : "☀️";
})();

themeToggle.addEventListener("click", () => {
  const isDark = document.body.classList.contains("dark-theme");
  document.body.className = isDark ? "light-theme" : "dark-theme";
  themeToggle.textContent = isDark ? "☀️" : "🌙";
  localStorage.setItem("kidde-theme", document.body.className);
});

/* ── File Upload ────────────────────────────────────────── */

// El botón "Seleccionar Archivo" usa la API correcta según el navegador
document.querySelector(".btn-outline").addEventListener("click", (e) => {
  e.stopPropagation();
  FS_API ? openWithFSAPI() : fileInput.click();
});

uploadZone.addEventListener("click", () => {
  FS_API ? openWithFSAPI() : fileInput.click();
});

uploadZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadZone.classList.add("drag-over");
});
uploadZone.addEventListener("dragleave", () => uploadZone.classList.remove("drag-over"));
uploadZone.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) readFileObject(file); // drag-drop no da FileHandle → fallback descarga
});

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) readFileObject(fileInput.files[0]);
});

/* File System Access API — abre Y guarda el handle para sobreescribir */
async function openWithFSAPI() {
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
      multiple: false,
    });
    state.fileHandle = handle;
    state.fileName   = handle.name;
    const file = await handle.getFile();
    await readFileObject(file);
    showFileStatus(`✅ Vinculado: ${handle.name} — se sobreescribirá automáticamente`, "ok");
  } catch (err) {
    if (err.name !== "AbortError") showFileStatus("❌ No se pudo abrir el archivo.", "error");
  }
}

/* Lee un File object y carga el workbook */
function readFileObject(file) {
  return new Promise((resolve) => {
    if (!file.name.endsWith(".xlsx")) {
      showFileStatus("❌ Solo se aceptan archivos .xlsx", "error");
      resolve();
      return;
    }
    if (!state.fileHandle) {
      // Vino de drag-drop o input clásico: guardamos nombre pero no handle
      state.fileName = file.name;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        state.workbook = XLSX.read(e.target.result, { type: "binary" });
        if (!state.fileHandle) {
          showFileStatus(`✅ Archivo cargado: ${file.name} (se descargará como nuevo archivo)`, "ok");
        }
      } catch {
        showFileStatus("❌ Error al leer el archivo Excel.", "error");
        state.workbook = null;
      }
      resolve();
    };
    reader.readAsBinaryString(file);
  });
}

function showFileStatus(msg, type) {
  fileStatus.textContent = msg;
  fileStatus.classList.remove("hidden");
  fileStatus.style.background = type === "ok" ? "" : "#f8d7da";
  fileStatus.style.color      = type === "ok" ? "" : "#721c24";
  fileStatus.style.borderColor= type === "ok" ? "" : "#f5c6cb";
}

/* ── Form Validation ────────────────────────────────────── */
const FIELDS = [
  { id: "empresa",      label: "Empresa" },
  { id: "matricula",    label: "Matrícula" },
  { id: "tipoLicencia", label: "Tipo de Licencia" },
  { id: "fechaIngreso", label: "Fecha de Ingreso" },
  { id: "horaIngreso",  label: "Hora de Ingreso" },
  { id: "fechaSalida",  label: "Fecha de Salida" },
  { id: "horaSalida",   label: "Hora de Salida" },
];

function validateForm() {
  let valid = true;
  FIELDS.forEach(({ id, label }) => {
    const el  = $(id);
    const err = $(`err-${id}`);
    const val = el.value.trim();
    el.classList.remove("invalid");
    err.textContent = "";

    if (!val) {
      el.classList.add("invalid");
      err.textContent = `${label} es obligatorio.`;
      valid = false;
    }
  });

  // Cross-date validation
  const fi = $("fechaIngreso").value;
  const fs = $("fechaSalida").value;
  if (fi && fs && fs < fi) {
    $("fechaSalida").classList.add("invalid");
    $("err-fechaSalida").textContent = "La fecha de salida no puede ser anterior a la de ingreso.";
    valid = false;
  }

  return valid;
}

/* ── Form Submit ────────────────────────────────────────── */
accessForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  clearMessage();
  if (!validateForm()) {
    showMessage("Por favor corrige los campos marcados.", "error");
    return;
  }

  const record = {
    empresa:      $("empresa").value.trim(),
    fechaIngreso: $("fechaIngreso").value,
    horaIngreso:  $("horaIngreso").value,
    matricula:    $("matricula").value.trim(),
    tipoLicencia: $("tipoLicencia").value,
    fechaSalida:  $("fechaSalida").value,
    horaSalida:   $("horaSalida").value,
  };

  // Prepend (most recent first in preview)
  state.records.unshift(record);
  renderTable();
  await exportAndDownload(record);
  resetForm();
});

/* ── Clear Button ───────────────────────────────────────── */
clearBtn.addEventListener("click", () => {
  resetForm();
  clearMessage();
});

function resetForm() {
  accessForm.reset();
  FIELDS.forEach(({ id }) => {
    $(id).classList.remove("invalid");
    $(`err-${id}`).textContent = "";
  });
}

/* ── Messages ───────────────────────────────────────────── */
function showMessage(msg, type) {
  formMessage.textContent = msg;
  formMessage.className   = `form-message ${type}`;
  formMessage.classList.remove("hidden");
  setTimeout(() => formMessage.classList.add("hidden"), 5000);
}
function clearMessage() {
  formMessage.classList.add("hidden");
}

/* ── Export Excel ───────────────────────────────────────── */
async function exportAndDownload(newRecord) {
  let wb = state.workbook
    ? cloneWorkbook(state.workbook)
    : createFreshWorkbook();

  const sheetName = wb.SheetNames[0];
  let ws = wb.Sheets[sheetName];

  // Convert sheet → array of arrays (AOA)
  let aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  // Ensure headers exist
  if (aoa.length === 0 || !rowIsHeader(aoa[0])) {
    aoa.unshift(HEADERS);
  }

  // Insert AFTER header row (index 1) → most recent at top
  aoa.splice(1, 0, recordToRow(newRecord));

  // Rebuild worksheet
  const newWs = XLSX.utils.aoa_to_sheet(aoa);
  styleWorksheet(newWs);
  wb.Sheets[sheetName] = newWs;

  // Persist updated workbook in state
  state.workbook = wb;

  // ── Intentar sobreescribir vía File System Access API ──
  if (state.fileHandle) {
    try {
      await writeToHandle(state.fileHandle, wb);
      showMessage(`✅ Archivo actualizado: ${state.fileHandle.name}`, "success");
      return;
    } catch (err) {
      if (err.name === "AbortError") {
        showMessage("⚠️ Guardado cancelado por el usuario.", "error");
        return;
      }
      // Si falla por cualquier razón, caemos al fallback
      console.warn("File System API falló, usando descarga:", err);
    }
  }

  // ── Fallback: descarga clásica ──
  const today = new Date().toISOString().split("T")[0];
  const name  = state.fileName
    ? state.fileName.replace(/\.xlsx$/i, `_${today}.xlsx`)
    : `Bitacora_${today}.xlsx`;
  XLSX.writeFile(wb, name);
}

/* Escribe el workbook al FileHandle (sobreescribe en disco) */
async function writeToHandle(handle, wb) {
  const wbBinary = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbBinary], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const writable = await handle.createWritable();
  await writable.write(blob);
  await writable.close();
}

function recordToRow(r) {
  return [
    r.empresa,
    r.fechaIngreso,
    r.horaIngreso,
    r.matricula,
    r.tipoLicencia,
    r.fechaSalida,
    r.horaSalida,
  ];
}

function rowIsHeader(row) {
  return row && row[0] === "Empresa";
}

function createFreshWorkbook() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([HEADERS]);
  XLSX.utils.book_append_sheet(wb, ws, "Bitácora");
  return wb;
}

function cloneWorkbook(wb) {
  // Deep clone via write + read
  const binary = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
  return XLSX.read(binary, { type: "binary" });
}

function styleWorksheet(ws) {
  // Column widths
  const colWidths = [28, 14, 12, 18, 15, 14, 12];
  ws["!cols"] = colWidths.map((w) => ({ wch: w }));
}

/* ── Manual Export Button ───────────────────────────────── */
exportBtn.addEventListener("click", async () => {
  if (!state.workbook) {
    showMessage("Agrega al menos un registro primero.", "error");
    return;
  }
  if (state.fileHandle) {
    try {
      await writeToHandle(state.fileHandle, state.workbook);
      showMessage(`✅ Archivo guardado: ${state.fileHandle.name}`, "success");
    } catch {
      showMessage("❌ No se pudo guardar el archivo.", "error");
    }
  } else {
    const today = new Date().toISOString().split("T")[0];
    const name  = state.fileName
      ? state.fileName.replace(/\.xlsx$/i, `_${today}.xlsx`)
      : `Bitacora_${today}.xlsx`;
    XLSX.writeFile(state.workbook, name);
  }
});

/* ── Render Table ───────────────────────────────────────── */
function renderTable() {
  const q = state.searchQuery.toLowerCase();
  const filtered = state.records.filter((r) =>
    Object.values(r).some((v) => v.toLowerCase().includes(q))
  );

  if (filtered.length === 0) {
    recordsBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="9">
          <div class="empty-state">
            <span class="empty-icon">${state.searchQuery ? "🔍" : "📋"}</span>
            <p>${state.searchQuery ? "No se encontraron registros con esa búsqueda." : "Sin registros aún. Llena el formulario para comenzar."}</p>
          </div>
        </td>
      </tr>`;
    recordCount.textContent = `${state.records.length} registro${state.records.length !== 1 ? "s" : ""}`;
    return;
  }

  recordsBody.innerHTML = filtered
    .map((r, i) => `
      <tr data-index="${state.records.indexOf(r)}">
        <td>${i + 1}</td>
        <td>${esc(r.empresa)}</td>
        <td>${esc(r.fechaIngreso)}</td>
        <td>${esc(r.horaIngreso)}</td>
        <td>${esc(r.matricula)}</td>
        <td><span class="badge">${esc(r.tipoLicencia)}</span></td>
        <td>${esc(r.fechaSalida)}</td>
        <td>${esc(r.horaSalida)}</td>
        <td>
          <button class="btn-danger" title="Eliminar registro" onclick="deleteRecord(${state.records.indexOf(r)})">🗑</button>
        </td>
      </tr>`)
    .join("");

  recordCount.textContent = `${filtered.length} registro${filtered.length !== 1 ? "s" : ""} ${state.searchQuery ? "(filtrados)" : ""}`;
}

function esc(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* ── Delete Record (preview only) ──────────────────────── */
window.deleteRecord = function (index) {
  if (!confirm("¿Eliminar este registro de la vista previa?")) return;
  state.records.splice(index, 1);
  renderTable();
};

/* ── Search ─────────────────────────────────────────────── */
searchInput.addEventListener("input", () => {
  state.searchQuery = searchInput.value.trim();
  renderTable();
});

/* ── Set default dates to today ─────────────────────────── */
(function prefillDates() {
  const today = new Date().toISOString().split("T")[0];
  $("fechaIngreso").value = today;
  $("fechaSalida").value  = today;

  const now = new Date();
  const hh  = String(now.getHours()).padStart(2, "0");
  const mm  = String(now.getMinutes()).padStart(2, "0");
  $("horaIngreso").value = `${hh}:${mm}`;
  $("horaSalida").value  = `${hh}:${mm}`;
})();
