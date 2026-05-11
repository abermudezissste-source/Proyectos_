/**************************
 * app.js - Dashboard EBDI
 **************************/

const mapa = document.getElementById("mapa");
const modal = document.getElementById("estadoModal");
const modalTitulo = document.getElementById("modalTitulo");
const modalTexto = document.getElementById("modalTexto");
const btnAceptar = document.getElementById("btnAceptar");
const btnDetalle = document.getElementById("btnDetalle");
const tooltip = document.getElementById("tooltipEstado");

let estadoSeleccionado = null;

mapa.addEventListener("load", () => {
  const svgDoc = mapa.contentDocument;
  if (!svgDoc) return;

  const estados = svgDoc.querySelectorAll("path");

  estados.forEach(estado => {

    estado.style.cursor = "pointer";

    // ✅ HOVER (ETIQUETA)
    estado.addEventListener("mouseenter", () => {
      tooltip.innerHTML = `
        <strong>${estado.dataset.estado || "Sin nombre"}</strong><br>
        ID: ${estado.id}
      `;
      tooltip.style.display = "block";
    });

    estado.addEventListener("mousemove", e => {
      tooltip.style.left = (e.clientX + 12) + "px";
      tooltip.style.top  = (e.clientY + 12) + "px";
    });

    estado.addEventListener("mouseleave", () => {
      tooltip.style.display = "none";
    });

    // ✅ CLICK (MODAL) — SE QUEDA IGUAL
    estado.addEventListener("click", () => {
      estadoSeleccionado = estado.id;

      modalTitulo.textContent = "Estado seleccionado";
      modalTexto.textContent =
        "ID: " + estado.id + "\n" +
        "Nombre: " + estado.dataset.estado;

      modal.style.display = "block";
    });

  });
});

btnAceptar.addEventListener("click", () => {
  modal.style.display = "none";
});

btnDetalle.addEventListener("click", () => {
  window.location.href =
    "matricula.html?estado=" + estadoSeleccionado;
});

window.addEventListener("load", () => {
  const mapa = document.getElementById("mapa");
  setTimeout(() => {
    mapa.classList.add("visible");
  }, 200);
});
