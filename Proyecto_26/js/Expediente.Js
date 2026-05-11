document.addEventListener("DOMContentLoaded", () => {
    console.log("✅ js/expediente.js cargado correctamente");

    // ===============================
    // 1. Parámetros de navegación
    // ===============================
    const params = new URLSearchParams(window.location.search);

    const matriculaId = params.get("matricula");
    const ebdiId = params.get("ebdi");
    const estadoId = params.get("estado");

    if (!matriculaId) {
        alert("No se recibió la Matrícula. No se puede abrir el expediente.");
        return;
    }

    // ===============================
    // 2. Encabezado del expediente
    // ===============================
    const titulo = document.getElementById("tituloExpediente");
    const lblMatricula = document.getElementById("Matricula");
    const lblEBDI = document.getElementById("EBDI");

    if (titulo) {
        titulo.textContent = `Expediente electrónico – Matrícula ${matriculaId}`;
    }

    if (lblMatricula) lblMatricula.textContent = matriculaId;
    if (lblEBDI) lblEBDI.textContent = ebdiId ?? "—";

    // ===============================
    // 3. Formulario
    // ===============================
    const form = document.getElementById("expedienteForm");
    if (!form) {
        console.error("❌ No existe el formulario expedienteForm");
        return;
    }

    // Prellenar campos fijos
    if (form.matriculaId) form.matriculaId.value = matriculaId;
    if (form.ebdiId) form.ebdiId.value = ebdiId ?? "";
    if (form.EstadoId) form.EstadoId.value = estadoId ?? "";

    // ===============================
    // 4. Modal Bootstrap (si existe)
    // ===============================
    const modalEl = document.getElementById("modalConfirmacion");
    let modalConfirmacion = null;

    if (modalEl && window.bootstrap) {
        modalConfirmacion = new bootstrap.Modal(modalEl);
    }

    // ===============================
    // 5. Cargar Base.xlsx
    // ===============================
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

        if (form.genero && r.Genero) form.genero.value = r.Genero;
        if (form.estrato && r.Estrato) form.estrato.value = r.Estrato;
    }

    // ===============================
    // 6. Inicialización
    // ===============================
    (async function initExpediente() {
        try {
            await cargarBase();
            const datosBase = obtenerDatosBaseMatricula(matriculaId);
            prellenarDesdeBase(datosBase);
        } catch (err) {
            console.error(err);
            alert("Error al cargar la información base del expediente");
        }
    })();

    // ===============================
    // 7. Guardar expediente
    // ===============================
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

            Nombres: form.nombres?.value ?? "",
            ApellidoPaterno: form.ap_paterno?.value ?? "",
            ApellidoMaterno: form.ap_materno?.value ?? "",
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

    // ===============================
    // 8. Volver a matrículas
    // ===============================
    const btnVolver = document.getElementById("btnVolver");
    if (btnVolver) {
        btnVolver.addEventListener("click", () => {
            const params = new URLSearchParams();
            if (estadoId) params.set("estado", estadoId);
            if (ebdiId) params.set("ebdi", ebdiId);

            window.location.href = `matricula.html${params.toString() ? "?" + params : ""}`;
        });
    }
});