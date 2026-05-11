// ===============================
// matricula.js — versión limpia
// ===============================
console.log("✅ matricula.js cargado correctamente");

// ---------- Config ----------
const EXCEL_URL = "Base.xlsx";

// ---------- DOM ----------
const titulo = document.getElementById("titulo");
const estadoLabel = document.getElementById("estadoLabel");
const status = document.getElementById("status");
const resultado = document.getElementById("resultado");
const pagerInfo = document.getElementById("pagerInfo");

const selTipo = document.getElementById("selTipo");
const selEBDI = document.getElementById("selEBDI");
const selEstrato = document.getElementById("selEstrato");
const selGenero = document.getElementById("selGenero");
const txtMatricula = document.getElementById("txtMatricula");

const btnMostrar = document.getElementById("btnMostrar");

const statEBDI = document.getElementById("statEBDI");
const statMatriculas = document.getElementById("statMatriculas");
const statRegistros = document.getElementById("statRegistros");

// ---------- Estado ----------
let DATA = [];          // datos ya filtrados por estado
let VIEW = [];          // datos filtrados por UI
let PAGE = 1;
const PAGE_SIZE = 25; // pares EBDI/Matrícula por página


// ---------- Utils ----------
const uniq = arr => [...new Set(arr.filter(v => v !== "" && v != null))];

function initSelect(sel, label = "Todos") {
  sel.innerHTML = `<option value="__ALL__">${label}</option>`;
}

function promedio(arr, campo) {
  if (!arr.length) return "0.0";
  const s = arr.reduce((a, r) => a + (+r[campo] || 0), 0);
  return (s / arr.length).toFixed(1);
}

function ultimoRegistro(arr) {
  if (!arr.length) return null;
  return [...arr].reduce((a, b) =>
    new Date(a["Fecha medición"]) > new Date(b["Fecha medición"]) ? a : b
  );
}

// ---------- Leer estado desde URL ----------
const params = new URLSearchParams(window.location.search);
const estado = params.get("estado");

if (!estado) {
  status.textContent = "❌ Falta parámetro ?estado=...";
  throw new Error("Estado no definido");
}

titulo.textContent = "Histórico de Matrículas por EBDI";
estadoLabel.textContent = estado;

// ---------- Cargar Excel ----------
async function cargarExcel() {
  status.textContent = "Cargando Base.xlsx…";
 
  const resp = await fetch(EXCEL_URL, { cache: "no-store" });
  if (!resp.ok) throw new Error("No se pudo cargar el Excel");
   iniciarProgreso();

  const ab = await resp.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
  
function obtenerRepresentacionEstado(data, estadoId) {
  if (!estadoId || !Array.isArray(data)) {
    return estadoId || "—";
  }

  const fila = data.find(r => r.EstadoId === estadoId);
  if (!fila) return estadoId;

  return (
    fila["Representación"] ||
    fila["Representacion"] ||
    fila["REPRESENTACION"] ||
    estadoId
  );
}

  
  // Normalizar
  DATA = rows
    .map(r => ({
      ...r,
      EstadoId: String(r.EstadoId).trim(),
      EBDI: String(r.EBDI).trim(),
      Matricula: String(r.Matricula).trim(),
      Tipo: String(r.Tipo).trim(),
      Estrato: String(r.Estrato).trim(),
      Genero: String(r.Genero).trim()
    }))
    .filter(r => r.EstadoId === estado);

  status.textContent = `Datos cargados (${DATA.length.toLocaleString("es-MX")} registros)`;
  finalizarProgreso();

  initFiltros();
}

// ---------- Inicializar filtros ----------
function initFiltros() {
  initSelect(selTipo);
  initSelect(selEBDI);
  initSelect(selEstrato);
  initSelect(selGenero);

  uniq(DATA.map(r => r.Tipo)).sort().forEach(v =>
    selTipo.insertAdjacentHTML("beforeend", `<option>${v}</option>`)
  );
  uniq(DATA.map(r => r.EBDI)).sort().forEach(v =>
    selEBDI.insertAdjacentHTML("beforeend", `<option>${v}</option>`)
  );
  uniq(DATA.map(r => r.Estrato)).sort().forEach(v =>
    selEstrato.insertAdjacentHTML("beforeend", `<option>${v}</option>`)
  );
  uniq(DATA.map(r => r.Genero)).sort().forEach(v =>
    selGenero.insertAdjacentHTML("beforeend", `<option>${v}</option>`)
  );
}

// ---------- Aplicar filtros ----------
function aplicarFiltros() {
  VIEW = DATA.filter(r => {
    if (selTipo.value !== "__ALL__" && r.Tipo !== selTipo.value) return false;
    if (selEBDI.value !== "__ALL__" && r.EBDI !== selEBDI.value) return false;
    if (selEstrato.value !== "__ALL__" && r.Estrato !== selEstrato.value) return false;
    if (selGenero.value !== "__ALL__" && r.Genero !== selGenero.value) return false;
    if (txtMatricula.value && r.Matricula !== txtMatricula.value.trim()) return false;
    return true;
  });
}

// ---------- Render ----------
function render() {
  aplicarFiltros();
  resultado.innerHTML = "";

  // ===============================
  // Estadísticas
  // ===============================
  statEBDI.textContent = uniq(VIEW.map(r => r.EBDI)).length;
  statMatriculas.textContent = uniq(VIEW.map(r => r.Matricula)).length;
  mostrarNumero("statRegistros", VIEW.length);
  // ===============================
  // Construir pares EBDI/Matrícula
  // ===============================
  const map = new Map();
  VIEW.forEach(r => {
    const key = `${r.EBDI}||${r.Matricula}`;
    if (!map.has(key)) {
      map.set(key, { EBDI: r.EBDI, Matricula: r.Matricula, registros: [] });
    }
    map.get(key).registros.push(r);
  });

  const pares = [...map.values()].sort((a, b) =>
    a.EBDI.localeCompare(b.EBDI) ||
    a.Matricula.localeCompare(b.Matricula)
  );

  const totalPares = pares.length;
  const totalPages = Math.max(1, Math.ceil(totalPares / PAGE_SIZE));
  PAGE = Math.min(PAGE, totalPages);

  const inicio = (PAGE - 1) * PAGE_SIZE;
  const visibles = pares.slice(inicio, inicio + PAGE_SIZE);

  // ===============================
  // Agrupar visibles por EBDI
  // ===============================
  const porEBDI = {};
  visibles.forEach(p => {
    porEBDI[p.EBDI] ??= [];
    porEBDI[p.EBDI].push(p);
  });

  // ===============================
  // Render visual
  // ===============================
  Object.entries(porEBDI).forEach(([ebdi, lista]) => {
    const detE = document.createElement("details");
    detE.open = true;
    detE.innerHTML = `<summary><b>EBDI:</b> ${ebdi}</summary>`;

    lista.forEach(p => {
  const detM = document.createElement("details");

  detM.innerHTML = `
    <summary class="summary-matricula">
      <span>
        <b>Matrícula:</b> ${p.Matricula}
        <span class="muted">(${p.registros.length})</span>
      </span>

      <button
        type="button"
        class="btn-expediente"
        onclick="irAExpediente('${p.Matricula}', '${p.EBDI}')">
        Ir al Expediente
      </button>
    </summary>

    ${crearTablaMatricula(p.registros)}
  `;

  detE.appendChild(detM);
});

    resultado.appendChild(detE);
  });

function crearTablaMatricula(registros) {
  return `
  <div style="overflow-x:auto; max-width:100%;">
  <table border="1" cellpadding="4" cellspacing="0" style="width:100%; border-collapse:collapse; margin-top:8px">
    <thead>
      <tr>
        <th>#</th>
        <th>Año</th>
        <th>Mes</th>
        <th>Fecha medición</th>
        <th>UR</th>
        <th>EBDI</th>
        <th>CentroNombre</th>
        <th>Genero</th>
        <th>Meses Cumplidos</th>
        <th>Peso</th>
        <th>D. Peso</th>
        <th>Talla</th>
        <th>D. Talla</th>
        <th>IMC</th>
        <th>D. IMC</th>
        <th>PC</th>
        <th>D. PC</th>
        <th>Clv.IMC</th>
        <th>Propuesta Indicador</th>
        <th>Tendencia</th>
      </tr>
    </thead>
    <tbody>
      ${registros.map((r, i) => `
        <tr>
          <td>${i + 1}</td>
          <td>${r["Año"] ?? "-"}</td>
          <td>${r["Mes"] ?? "-"}</td>
          <td>${r["Fecha medición"] ?? "-"}</td>
          <td>${r["UR"] ?? "-"}</td>
          <td>${r["EBDI"] ?? "-"}</td>
          <td>${r["CentroNombre"] ?? "-"}</td>
          <td>${r["Genero"] ?? "-"}</td>
          <td>${r["Meses Cumplidos"] ?? "-"}</td>
          <td>${r["Peso"] ?? "-"}</td>
          <td>${r["D. Peso"] ?? "-"}</td>
          <td>${r["Talla"] ?? "-"}</td>
          <td>${r["D. Talla"] ?? "-"}</td>
          <td>${r["IMC"] ?? "-"}</td>
          <td>${r["D. IMC"] ?? "-"}</td>
          <td>${r["PC"] ?? "-"}</td>
          <td>${r["D. PC"] ?? "-"}</td>
          <td>${r["Clv.IMC"] ?? "-"}</td>
          <td>${r["Propuesta Indicador"] ?? "-"}</td>
          <td>${r["Tendencia"] ?? "-"}</td>

<td data-imc="${r["D. IMC"] ?? ""}">
  ${r["D. IMC"] ?? "-"}
</td>
        </tr>
      `).join("")}

    </tbody>
  </table>
  `;
}

  // ===============================
  // Info de paginación
  // ===============================
  pagerInfo.textContent =
    `Página ${PAGE} / ${totalPages} — Mostrando ${visibles.length} de ${totalPares} pares (EBDI/Matrícula)`;
}

// ---------- Eventos ----------
btnMostrar.addEventListener("click", () => {
  render();
});

// ---------- Init ----------
cargarExcel().catch(err => {
  console.error(err);
  status.textContent = "❌ Error al cargar datos";
});

// --------Botones--------
document.getElementById("btnPrev").addEventListener("click", () => {
  if (PAGE > 1) {
    PAGE--;
    render();
  }
});

document.getElementById("btnNext").addEventListener("click", () => {
  PAGE++;
  render();
});


document.getElementById("btnVolver").addEventListener("click", () => {
  window.location.href = "index.html";
});

let EXPANDIDO = true;

document.getElementById("btnExpandir").addEventListener("click", () => {
  const detalles = resultado.querySelectorAll("details");
  detalles.forEach(d => d.open = !EXPANDIDO);

  EXPANDIDO = !EXPANDIDO;
  document.getElementById("btnExpandir").textContent =
    EXPANDIDO ? "⤵️ Expandir todo" : "⤴️ Colapsar todo";
});

document.getElementById("btnExportEBDI").addEventListener("click", () => {
  if (!VIEW.length) {
    alert("No hay datos para exportar");
    return;
  }

  const wb = XLSX.utils.book_new();

  const porEBDI = {};
  VIEW.forEach(r => {
    porEBDI[r.EBDI] ??= [];
    porEBDI[r.EBDI].push(r);
  });

  Object.entries(porEBDI).forEach(([ebdi, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, `EBDI_${ebdi}`.substring(0, 31));
  });

  XLSX.writeFile(wb, `Historico_Matriculas_${estado}.xlsx`);
});

function mostrarNumero(id, valor) {
  document.getElementById(id).textContent =
    Number(valor).toLocaleString("es-MX");
}

function irAExpediente(matricula, ebdi) {
  window.location.href =
    `Expediente.html?matricula=${encodeURIComponent(matricula)}&ebdi=${encodeURIComponent(ebdi)}`;
}


const progressBar = document.getElementById("progressBar");
const loader = document.getElementById("loader");

let fakeProgressTimer = null;

function iniciarProgreso() {
  let progreso = 5;
  progressBar.style.width = progreso + "%";

  // MOSTRAR GIF
  loader.style.display = "block";

  fakeProgressTimer = setInterval(() => {
    if (progreso < 90) {
      progreso += Math.random() * 10;
      progressBar.style.width = Math.min(progreso, 90) + "%";
    }
  }, 300);
}

function finalizarProgreso() {
  if (fakeProgressTimer) {
    clearInterval(fakeProgressTimer);
  }

  progressBar.style.width = "100%";

  // OCULTAR GIF después de completar
  setTimeout(() => {
    loader.style.display = "none";
  }, 300);
}

