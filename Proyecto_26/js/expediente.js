// console.log("Genero Excel:", r.Genero);

document.addEventListener("DOMContentLoaded", () => {
    console.log("✅ js/expediente.js cargado correctamente");
    document.body.classList.add("fade-in");

    // 1. Parámetros de navegación
    const params = new URLSearchParams(window.location.search);

    const matriculaId = params.get("matricula");
    const ebdiId = params.get("ebdi");
    const estadoId = params.get("estado");

    if (!matriculaId) {
        alert("No se recibió la Matrícula. No se puede abrir el expediente.");
        return;
    }

    // 2. Encabezado del expediente
    const titulo = document.getElementById("tituloExpediente");
    const lblMatricula = document.getElementById("Matricula");
    const lblEBDI = document.getElementById("EBDI");

    if (titulo) {
        titulo.textContent = `Expediente Electrónico – Matrícula ${matriculaId}`;
    }

    if (lblMatricula) lblMatricula.textContent = matriculaId;
    if (lblEBDI) lblEBDI.textContent = ebdiId ?? "—";

    // 3. Formulario
    const form = document.getElementById("expedienteForm");
    if (!form) {
        console.error("❌ No existe el formulario expedienteForm");
        return;
    }

    // Prellenar campos fijos
    if (form.matriculaId) form.matriculaId.value = matriculaId;
    if (form.ebdiId) form.ebdiId.value = ebdiId ?? "";
    if (form.EstadoId) form.EstadoId.value = estadoId ?? "";

    // 4. Modal Bootstrap (si existe)
    const modalEl = document.getElementById("modalConfirmacion");
    let modalConfirmacion = null;

    if (modalEl && window.bootstrap) {
        modalConfirmacion = new bootstrap.Modal(modalEl);
    }

    // 5. Cargar Base.xlsx
    let BASE_DATA = [];

    async function cargarBase() {
        const resp = await fetch("Base.xlsx", { cache: "no-store" });
        if (!resp.ok) throw new Error("No se pudo cargar Base.xlsx");

        const ab = await resp.arrayBuffer();
        const wb = XLSX.read(ab, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];

        BASE_DATA = XLSX.utils.sheet_to_json(ws, {
            defval: "",
            raw: false
        });
    }

    function obtenerDatosBaseMatricula(matricula) {
        return BASE_DATA.filter(r => String(r.Matricula) === String(matricula));
    }

function prellenarDesdeBase(registros) {
    if (!registros.length) return;

    const r = registros[registros.length - 1];

    // MAPEO PERSONALIZADO
    if (form.estrato && r.Estrato) {
        form.estrato.value = r.Estrato;
    }

    if (form.curp && r.CURP) {
        form.curp.value = r.CURP;
    }

    if (form["Nombre completo"] && r["Nombre completo"]) {
        form["Nombre completo"].value = r["Nombre completo"];
    }

    if (form.matriculaId && r.Matricula) {
        form.matriculaId.value = r.Matricula;
    }

    if (form.ebdiId && r.EBDI) {
        form.ebdiId.value = r.EBDI;
    }

    if (form.EstadoId && r.EstadoId) {
        form.EstadoId.value = r.EstadoId;
    }

    if (form["Fecha de Nacimiento"] && r["Fecha de Nacimiento"]) {
        form["Fecha de Nacimiento"].value = r["Fecha de Nacimiento"];
    }

// FECHA + EDAD COMPLETA
if (form.fechaNacimiento && r["Fecha de Nacimiento"]) {

    const fechaOriginal = r["Fecha de Nacimiento"];
    console.log("Fecha original:", fechaOriginal);

    const fechaConvertida = convertirFecha(fechaOriginal);
    console.log("Fecha convertida:", fechaConvertida);

    if (fechaConvertida) {

        // asignar fecha al input
        form.fechaNacimiento.value = fechaConvertida;

        // calcular edad con meses
        const edadCompleta = calcularEdadMeses(fechaConvertida);
        console.log("Edad completa:", edadCompleta);

        if (form.edad) {
            form.edad.value = edadCompleta;
        }

    } else {
        console.warn("No se pudo convertir la fecha:", fechaOriginal);
        form.fechaNacimiento.value = "";
    }
}

// CONVERTIR FECHA (AUTO)
function convertirFecha(fecha) {
    if (!fecha) return "";

    if (typeof fecha === "string" && fecha.includes("/")) {

        const partes = fecha.split("/");
        if (partes.length !== 3) return "";

        let p1 = partes[0];
        let p2 = partes[1];
        let p3 = partes[2];

        // año corto → largo
        if (p3.length === 2) {
            p3 = "20" + p3;
        }

        // detectar formato
        if (parseInt(p1) > 12) {
            // dd/mm/yyyy
            return `${p3}-${p2.padStart(2, "0")}-${p1.padStart(2, "0")}`;
        } else {
            // mm/dd/yyyy
            return `${p3}-${p1.padStart(2, "0")}-${p2.padStart(2, "0")}`;
        }
    }

    return "";
}

// EDAD AÑOS + MESES
function calcularEdadMeses(fechaNacimiento) {

    if (!fechaNacimiento) return "";

    const hoy = new Date();
    const nacimiento = new Date(fechaNacimiento);

    if (isNaN(nacimiento)) return "";

    let años = hoy.getFullYear() - nacimiento.getFullYear();
    let meses = hoy.getMonth() - nacimiento.getMonth();
    let dias = hoy.getDate() - nacimiento.getDate();

    if (dias < 0) {
        meses--;
    }

    if (meses < 0) {
        años--;
        meses += 12;
    }
    return `${años} años, ${meses} meses`;
}

// GÉNERO (transformación)
if (form.genero && r.Genero) {

    const genero = String(r.Genero).trim().toUpperCase();

    if (genero === "M") {
        form.genero.value = "Masculino";
    } else if (genero === "F") {
        form.genero.value = "Femenino";
    } else {
        form.genero.value = "";
    }
}
}

// 6. Inicialización
(async function initExpediente() {

    function ocultarLoader() {
        const loader = document.getElementById("loaderExpediente");
        if (loader) {
            loader.style.display = "none";
        }
    }

    try {
        await cargarBase();
        const datosBase = obtenerDatosBaseMatricula(matriculaId);
        prellenarDesdeBase(datosBase);

    } catch (err) {
        console.error(err);
        alert("Error al cargar la información base del expediente");

    } finally {
        //  SIEMPRE se oculta, haya error o no
        ocultarLoader();
    }

    // 7. Guardar expediente
    const btnGuardar = document.getElementById("btnGuardar");
    if (btnGuardar) {
        btnGuardar.addEventListener("click", guardarExpediente);
    }

    function guardarExpediente() {
        const registro = {
            Matricula: matriculaId,
            EstadoId: estadoId,
            EBDI: ebdiId,
            Fecha: new Date().toISOString().slice(0, 10),

            CURP: form.curp?.value ?? "",
            Genero: form.genero?.value ?? "",
            Estrato: form.estrato?.value ?? "",

            Telefono: form.telefono?.value ?? "",
            Whatsapp: form.whatsapp?.value ?? "",
            Correo: form.correo?.value ?? "",
            Direccion: form.direccion?.value ?? "",

            AntecedentesPatologicos: form.ant_personales_patologicos?.value ?? "",
            AntecedentesNoPatologicos: form.ant_personales_no_patologicos?.value ?? "",
            Alergias: form.alergias?.value ?? "",
            Observaciones: form.observaciones?.value ?? ""
        };

        const ws = XLSX.utils.json_to_sheet([registro]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Expediente");

        XLSX.writeFile(wb, `Expediente_${matriculaId}.xlsx`);
    }

    // 8. Volver
    const btnVolver = document.getElementById("btnVolver");

    if (btnVolver) {
        btnVolver.addEventListener("click", () => {
            document.body.classList.remove("fade-in");

            setTimeout(() => {
                window.location.href = `matricula.html?ebdi=${ebdiId}&estado=${estadoId}`;
            }, 300);
        });
    }

})();

    // 7. Guardar expediente
    const btnGuardar = document.getElementById("btnGuardar");
    if (btnGuardar) {
        btnGuardar.addEventListener("click", guardarExpediente);
    }

    function guardarExpediente() {
        const registro = {
            Matricula: matriculaId,
            EstadoId: estadoId,
            EBDI: ebdiId,
            Fecha: new Date().toISOString().slice(0, 10),

            CURP: form.curp?.value ?? "",
            Genero: form.genero?.value ?? "",
            Estrato: form.estrato?.value ?? "",

            Telefono: form.telefono?.value ?? "",
            Whatsapp: form.whatsapp?.value ?? "",
            Correo: form.correo?.value ?? "",
            Direccion: form.direccion?.value ?? "",

            AntecedentesPatologicos: form.ant_personales_patologicos?.value ?? "",
            AntecedentesNoPatologicos: form.ant_personales_no_patologicos?.value ?? "",
            Alergias: form.alergias?.value ?? "",
            Observaciones: form.observaciones?.value ?? ""
        };

        const ws = XLSX.utils.json_to_sheet([registro]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Expediente");

        XLSX.writeFile(wb, `Expediente_${matriculaId}.xlsx`);
    }

    // 8. Volver a matrículas
   const btnVolver = document.getElementById("btnVolver");

if (btnVolver) {

btnVolver.addEventListener("click", () => {
  document.body.classList.remove("fade-in");

  setTimeout(() => {
    window.location.href = `matricula.html?ebdi=${ebdiId}&estado=${estadoId}`;
  }, 300);
});

}
});

// 9. Modulo de Seguridad
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
