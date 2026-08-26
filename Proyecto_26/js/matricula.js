// matricula.js — versión limpia
console.log("✅ matricula.js cargado correctamente");

document.addEventListener("DOMContentLoaded", () => {
  document.body.classList.add("fade-in");
});

// ---------- Config ----------
const EXCEL_URL = "Base.xlsx";

// ---------- DOM ----------
const titulo = document.getElementById("titulo");
const estadoLabel = document.getElementById("estadoLabel");
const status = document.getElementById("status");
const resultado = document.getElementById("resultado");
const pagerInfo = document.getElementById("pagerInfo");

const filtroRegion = document.getElementById("filtroRegion");
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
// Charts globales
let primeraCarga = true;
let chartEstado = null;
let TOTAL_GLOBAL = 0;
// let DATA_GLOBAL = [];
let chartEBDI = null;
let chartTendencia = null;
let chartPeso = null;
let chartTalla = null;
let chartGenero = null;

Chart.register({
  id: 'labelsEstado',
  afterDatasetsDraw(chart) {
    if (chart.canvas.id !== "graficaEstado") return;

    const { ctx } = chart;

    chart.data.datasets.forEach((dataset, i) => {
      const meta = chart.getDatasetMeta(i);

      meta.data.forEach((bar, index) => {
        const valor = dataset.data[index];

        const porcentaje = ((valor / TOTAL_GLOBAL) * 100).toFixed(1);

        ctx.fillStyle = "#000";
        ctx.font = "12px Arial";
        ctx.textAlign = "center";

        ctx.fillText(
          porcentaje + "%",
          bar.x,
          bar.y - 5
        );
      });
    });
  }
});


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
const ebdi = params.get("ebdi");

if (!estado) {
  status.textContent = "❌ Falta parámetro ?estado=...";
  throw new Error("Estado no definido");
}

titulo.textContent = "Histórico de Matrículas en el Ciclo de Servicio por EBDI";
estadoLabel.textContent = estado;

let progressBar = document.getElementById("progressBar");
let loader = document.getElementById("loader");
let fakeProgressTimer = null;

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
  TOTAL_GLOBAL = rows.length;  //Total Nacional
  // DATA_GLOBAL = rows;
    
  // Normalizar
  DATA = rows
    .map(r => ({
      ...r,
      EstadoId: String(r.EstadoId).trim(),
      EBDI: String(r.EBDI).trim(),
      Matricula: String(r.Matricula).trim(),
      Tipo: String(r.Tipo).trim(),
      Estrato: String(r.Estrato).trim(),
      Genero: String(r.Genero).trim(),
     
      //Regionales desde Representación
    Region: normalizarTexto(r["Representación"])
    }))
    .filter(r => r.EstadoId === estado);

  status.textContent = `Datos cargados (${DATA.length.toLocaleString("es-MX")} registros)`;
  finalizarProgreso();

// Configuración de filtros para Regionales
configurarFiltroRegion(estado, DATA);
initFiltros();
// Auto despliegue inicial
render();
}

//Guardando DATA GLOBAL
localStorage.setItem("DATA", JSON.stringify(DATA));

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

function normalizarTexto(txt) {
  return String(txt || "")
    .trim()
    .toUpperCase();
}

// ---------- Aplicar filtros ----------
function aplicarFiltros() {

  let region = filtroRegion ? filtroRegion.value : "Todas";

  VIEW = DATA.filter(r => {

    if (region !== "Todas" && r.Region !== normalizarTexto(region)) return false;
    if (selTipo.value !== "__ALL__" && r.Tipo !== selTipo.value) return false;
    if (selEBDI.value !== "__ALL__" && r.EBDI !== selEBDI.value) return false;
    if (selEstrato.value !== "__ALL__" && r.Estrato !== selEstrato.value) return false;
    if (selGenero.value !== "__ALL__" && r.Genero !== selGenero.value) return false;
    if (txtMatricula.value && r.Matricula !== txtMatricula.value.trim()) return false;

    return true;
  });
}

// Filtrando Regiones
function configurarFiltroRegion(estado, DATA) {
  const contenedor = document.getElementById("contenedorRegion");
  const select = document.getElementById("filtroRegion");

  if (!contenedor || !select) return;
  //Limpiar
  select.innerHTML = '<option value="Todas">Todas</option>';
  //Solo para CDMX
  if (estado === "MXCMX") {
    contenedor.style.display = "inline-block";
  //Obtener Representaciones Regionales
    const regionesUnicas = [...new Set(DATA.map(d => d.Region).filter(Boolean))];

console.log("Representaciones detectadas:", regionesUnicas);
console.log("✅ VIEW:", VIEW.length);
console.log("✅ TOTAL_GLOBAL:", TOTAL_GLOBAL);


    regionesUnicas.sort();

    regionesUnicas.forEach(region => {
      const option = document.createElement("option");

      option.value = region;

      // Estética quitar mayúsculas)
      option.textContent = region
        .toLowerCase()
        .replace(/\b\w/g, l => l.toUpperCase());

      select.appendChild(option);
    });

  } else {
    contenedor.style.display = "none";
  }
}

if (filtroRegion) {
  filtroRegion.addEventListener("change", () => {
    PAGE = 1;
    render();
  });
}

function agruparPor(arr, campo) {
  const mapa = {};

  arr.forEach(item => {
    const key = item[campo] || "SIN DATO";
    mapa[key] = (mapa[key] || 0) + 1;
  });
  return mapa;
}

function renderGraficas(data) {

  if (!data || !data.length) return;

  const porEstado = agruparPor(data, "Estado");
  const porEBDI = agruparPor(data, "EBDI");
  const porTendencia = agruparPor(
  data.filter(r =>
    r["D. IMC"] &&
    r["D. IMC"] !== "-" &&
    r["D. IMC"] !== "Sin medición"
  ),
  "D. IMC"
  );
    const porPeso = agruparPor(
  data.filter(r =>
    r["D. Peso"] &&
    r["D. Peso"] !== "-" &&
    r["D. Peso"] !== "Sin medición"
  ),
  "D. Peso"
  );
    const porTalla = agruparPor(
  data.filter(r =>
    r["D. Talla"] &&
    r["D. Talla"] !== "-" &&
    r["D. Talla"] !== "Sin medición"
  ),
  "D. Talla"
  );
  const porGenero = agruparPor(data, "Genero");

Chart.defaults.font.size = 13;
Chart.defaults.font.family = "Arial";
Chart.defaults.color = "#333";

  // ========= ESTADO =========
const ctxEstado = document.getElementById("graficaEstado");

if (ctxEstado) {

  const labelsEstado = Object.keys(porEstado);
  const valoresEstado = Object.values(porEstado);

  const totalEstado = valoresEstado.reduce((a, b) => a + b, 0);
  const porcentajeEstado = ((totalEstado / TOTAL_GLOBAL) * 100).toFixed(2);

  if (!chartEstado) {
    chartEstado = new Chart(ctxEstado, {
      type: "bar",
      data: {
        labels: labelsEstado,
        datasets: [{
          label: "% del total nacional",
          data: valoresEstado,
          backgroundColor: "#680903"
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(context) {
                return `${totalEstado.toLocaleString()} registros (${porcentajeEstado}%) del total nacional (${TOTAL_GLOBAL.toLocaleString()})`;
              }
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true
          }
        }
      }
    });

  } else {
    chartEstado.data.labels = labelsEstado;
    chartEstado.data.datasets[0].data = valoresEstado;
    chartEstado.update(primeraCarga ? undefined : "none");
  }
}

  // ========= EBDI =========
  const ctxEBDI = document.getElementById("graficaEBDI");

if (ctxEBDI) {

  const entriesEBDI = Object.entries(porEBDI)
    .sort((a, b) => b[1] - a[1]);

  const labelsEBDI = entriesEBDI.map(e => e[0]);
  const valoresEBDI = entriesEBDI.map(e => e[1]);

  if (!chartEBDI) {
    chartEBDI = new Chart(ctxEBDI, {
      type: "bar",
      data: {
        labels: labelsEBDI,
        datasets: [{
          label: "EBDI",
          data: valoresEBDI,
          backgroundColor: "#e6e4d3"
        }]
      },
      options: {
        indexAxis: 'y', // ✅ horizontal
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(context) {
                return context.raw + " registros";
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true
          },
          y: {
            ticks: {
              autoSkip: false // ✅ muestra todos los EBDI
            }
          }
        }
      }
    });

  } else {
    chartEBDI.data.labels = labelsEBDI;
    chartEBDI.data.datasets[0].data = valoresEBDI;
    chartEBDI.update("none");
  }
}

  // ========= TENDENCIA =========
  const ctxT = document.getElementById("graficaTendencia");

  if (ctxT) {

    const labelsT = Object.keys(porTendencia);
    const valoresT = Object.values(porTendencia);
    const coloresT = labelsT.map(label => {
  const l = label.toLowerCase();

  if (l.includes("normal")) return "#4caf50";     // verde
  if (l.includes("posible riesgo de sobrepeso")) return "#eeff00"; // amarillo
  if (l.includes("sobrepeso")) return "#ff9800"; // naranja
  if (l.includes("obesidad")) return "#f44336";  // rojo

  return "#9e9e9e"; // gris para bajo peso
});

    if (!chartTendencia) {
      chartTendencia = new Chart(ctxT, {
        type: "bar",
        data: {
          labels: labelsT,
          
        datasets: [{
        data: valoresT,
        backgroundColor: coloresT
        }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          }
        }
      });

    } else {
      chartTendencia.data.labels = labelsT;
      chartTendencia.data.datasets[0].data = valoresT;
      chartTendencia.update(primeraCarga ? undefined : "none");
    }
  }

  // ========= PESO =========
  const ctxPeso = document.getElementById("graficaPeso");

  if (ctxPeso) {

    const labelsPeso = Object.keys(porPeso);
    const valoresPeso = Object.values(porPeso);
    const coloresPeso = labelsPeso.map(label => {
    const l = label.toLowerCase();

    if (l.includes("normal")) return "#4caf50";     // verde
    if (l.includes("peso alto para su edad")) return "#eeff00"; // amarillo
    if (l.includes("peso bajo")) return "#ff9800"; // naranja
    if (l.includes("peso bajo severo")) return "#f44336";  // rojo

   return "#9e9e9e"; // gris para alguna otra categoría
});

    if (!chartPeso) {
      chartPeso = new Chart(ctxPeso, {
        type: "bar",
        data: {
          labels: labelsPeso,
          
        datasets: [{
        data: valoresPeso,
        backgroundColor: coloresPeso
        }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          }
        }
      });

    } else {
      chartPeso.data.labels = labelsPeso;
      chartPeso.data.datasets[0].data = valoresPeso;
      chartPeso.update(primeraCarga ? undefined : "none");
    }
  }

    // ========= TALLA =========
  const ctxTalla = document.getElementById("graficaTalla");

  if (ctxTalla) {

    const labelsTalla = Object.keys(porTalla);
    const valoresTalla = Object.values(porTalla);
    const coloresTalla = labelsTalla.map(label => {
    const l = label.toLowerCase();

     if (l.includes("normal")) return "#4caf50";     // verde
     if (l.includes("talla baja")) return "#eeff00"; // amarillo
     if (l.includes("talla muy alta")) return "#ff9800"; // naranja
     if (l.includes("talla baja severa")) return "#f44336";  // rojo

   return "#9e9e9e"; // gris para alguna otra categoría
});

    if (!chartTalla) {
      chartTalla = new Chart(ctxTalla, {
        type: "bar",
        data: {
          labels: labelsTalla,
          
        datasets: [{
        data: valoresTalla,
        backgroundColor: coloresTalla
        }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false }
          }
        }
      });

    } else {
      chartTalla.data.labels = labelsTalla;
      chartTalla.data.datasets[0].data = valoresTalla;
      chartTalla.update(primeraCarga ? undefined : "none");
    }
  }

  // ========= GENERO =========
const ctxG = document.getElementById("graficaGenero");

if (ctxG) {

const entriesGenero = Object.entries(porGenero)
  .filter(([key]) => {
    const val = (key || "").toUpperCase();
    return val === "M" || val === "F";  // ✅ solo M y F
  });

const labelsG = entriesGenero.map(e => e[0]);
const valoresG = entriesGenero.map(e => e[1]);

const coloresGenero = labelsG.map(l => {
  const val = l.toUpperCase();

  if (val === "M") return "#42a5f5";
  if (val === "F") return "#ef5350";
});

  const totalG = valoresG.reduce((a, b) => a + b, 0);

  if (!chartGenero) {
    chartGenero = new Chart(ctxG, {
      type: "bar",
      data: {
        labels: labelsG,
        datasets: [{
          label: "Género",
          data: valoresG,
          backgroundColor: coloresGenero
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(context) {
                const valor = context.raw;
                const porcentaje = ((valor / totalG) * 100).toFixed(1);
                return `${valor} (${porcentaje}%)`;
              }
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true
          }
        }
      }
    });

  } else {
    chartGenero.data.labels = labelsG;
    chartGenero.data.datasets[0].data = valoresG;
    chartGenero.data.datasets[0].backgroundColor = coloresGenero;
    chartGenero.update(primeraCarga ? undefined : "none");
  }
    } 
  
primeraCarga = false;
}

Chart.register({
  id: 'labelsValores',
  afterDatasetsDraw(chart) {

    const { ctx } = chart;

    chart.data.datasets.forEach((dataset, i) => {

      const meta = chart.getDatasetMeta(i);

      // ✅ TOTAL para porcentajes
      const total = dataset.data.reduce((a, b) => a + Number(b || 0), 0);

      meta.data.forEach((element, index) => {

        const valor = Number(dataset.data[index]) || 0;

        if (valor === 0) return;

        const porcentaje = total > 0
          ? ((valor / total) * 100).toFixed(1)
          : 0;

        let texto = `${valor} (${porcentaje}%)`;

        ctx.fillStyle = "#333";
        ctx.font = "11px Arial";

        // ✅ DETECTAR tipo de gráfica
        const isHorizontal = chart.options.indexAxis === 'y';

        if (isHorizontal) {
          // 👉 barras horizontales (EBDI)
          ctx.textAlign = "left";
          ctx.fillText(texto, element.x + 10, element.y + 3);

        } else {
          // 👉 barras verticales
          ctx.textAlign = "center";
          ctx.fillText(texto, element.x, element.y - 10);
        }

      });

    });

  }
});

setTimeout(() => {
  renderGraficas(VIEW);
}, 100);

setTimeout(() => {
  chartEstado?.resize();
  chartEBDI?.resize();
  chartTendencia?.resize();
  chartGenero?.resize();
  chartPeso?.resize();
}, 200);

// ---------- Render ----------
function render() {
  aplicarFiltros();
  resultado.innerHTML = "";
  renderGraficas(VIEW);

   // Estadísticas
  statEBDI.textContent = uniq(VIEW.map(r => r.EBDI)).length;
  statMatriculas.textContent = uniq(VIEW.map(r => r.Matricula)).length;
  mostrarNumero("statRegistros", VIEW.length);
 
  // Construir pares EBDI/Matrícula
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
 
  // Agrupar visibles por EBDI
   const porEBDI = {};
  visibles.forEach(p => {
    porEBDI[p.EBDI] ??= [];
    porEBDI[p.EBDI].push(p);
  });

   // Render visual
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
        onclick="irAExpediente('${p.Matricula}', '${p.EBDI}', '${estado}')">
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

   // Info de paginación
    pagerInfo.textContent =
    `Página ${PAGE} / ${totalPages} — Mostrando ${visibles.length} de ${totalPares} pares (EBDI/Matrícula)`;
}

// ---------- Eventos ----------

// Debounce
function debounce(fn, delay = 300) {
  let timer;

  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

// Render con debounce
const renderDebounced = debounce(() => {
  PAGE = 1;
  render();
}, 300);

// Botón manual (por si acaso)
btnMostrar.addEventListener("click", () => {
  PAGE = 1;
  render();
});

// ✅ Asegurar que los elementos existen y agregar eventos
setTimeout(() => {

  [filtroRegion, selTipo, selEBDI, selEstrato, selGenero].forEach(sel => {
    if (!sel) {
      console.warn("Elemento no encontrado:", sel);
      return;
    }

    sel.addEventListener("change", () => {
      console.log("Cambio detectado en filtro");
      renderDebounced();
    });
  });

  if (txtMatricula) {
    txtMatricula.addEventListener("input", () => {
      console.log("Escribiendo matrícula...");
      renderDebounced();
    });
  }

}, 500);

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

// Abre Expediente de Matrícula
function irAExpediente(matricula, ebdi, estado) {

  const rol = localStorage.getItem("rol");

  // ⚠️ Solo bloquea si SÍ hay rol pero no es permitido
  if (rol && rol !== "admin") {
    alert("Acceso restringido. Favor de comunicarse con el administrador.");
    return;
  }

  // ✅ Si no hay rol, deja pasar (para no romper flujo actual)
  // ✅ Si es admin, también pasa
  document.body.classList.remove("fade-in");
  setTimeout(() => {
    window.location.href = `expediente.html?matricula=${matricula}&ebdi=${ebdi}&estado=${estado}`;
  }, 300);
}

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

// Modulo de Seguridad
window.addEventListener("DOMContentLoaded", () => {
    const usuario = localStorage.getItem("usuario");
    const rol = localStorage.getItem("rol");

    if (!usuario) {
        window.location.href = "login.html";
        return;
    }

    const userSpan = document.getElementById("usuarioLogeado");
    if (userSpan) {
        userSpan.textContent = usuario;
    }

    const rolSpan = document.getElementById("rolLogeado");
    if (rolSpan) {
        rolSpan.textContent = rol;
    }
});

function logout() {
    localStorage.clear();
    window.location.href = "login.html";
}

function verGraficas() {

  if (!DATA || !DATA.length) {
    alert("No hay datos para mostrar");
    return;
  }
}

function abrirModalGraficas() {

  const contenedor = document.getElementById("contenedorGraficas");
  const modalContenedor = document.getElementById("resumenGraficoModal");

  if (!contenedor || !modalContenedor) return;

  // ✅ SOLO mover si NO está ya en el modal
  if (!modalContenedor.contains(contenedor)) {
    modalContenedor.appendChild(contenedor);
  }

  const modal = document.getElementById("modalGraficas");
  modal.style.display = "block";
setTimeout(() => modal.classList.add("show"), 10);
}

async function exportarPDF() {

  const contenedor = document.getElementById("contenedorGraficas");

  if (!contenedor) {
    console.error("No existe contenedorGraficas");
    return;
  }

  const canvas = await html2canvas(contenedor, {
    scale: 2
  });

  const imgData = canvas.toDataURL("image/png");
  const pdf = new jspdf.jsPDF("p", "mm", "letter");
  const width = 210;
  const height = (canvas.height * width) / canvas.width;

  pdf.addImage(imgData, "PNG", 0, 10, width, height);
  pdf.save("graficas_dashboard.pdf");
}

function cerrarModal() {
  const modal = document.getElementById("modalGraficas");

  modal.classList.remove("show");

  setTimeout(() => {
    modal.style.display = "none";
  }, 300);
}

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    cerrarModal();
  }
});

document.getElementById("modalGraficas").addEventListener("click", (e) => {

  if (e.target.id === "modalGraficas") {
    cerrarModal();
  }

});