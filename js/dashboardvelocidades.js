let datosGlobales = [];
let miGraficoBarras = null;
let miGraficoTramos = null;
let miGraficoPendTramo = null;
let miGraficoPendHora = null;

document.getElementById('excelFile').addEventListener('change', handleFile, false);
document.getElementById('selectFecha').addEventListener('change', callbackFiltroFecha, false);
document.getElementById('selectTurno').addEventListener('change', procesarDatosPantalla, false);

function limpiarFiltros() {
    document.getElementById('excelFile').value = "";
    const selectF = document.getElementById('selectFecha');
    selectF.innerHTML = '<option value="">Seleccione un día...</option>';
    selectF.disabled = true;
    const selectT = document.getElementById('selectTurno');
    selectT.value = "";
    selectT.disabled = true;
    document.getElementById('panelCuenta').style.display = 'none';
    document.getElementById('statsRow').style.display = 'none';
    document.getElementById('dataRow').style.display = 'none';
    datosGlobales = [];
}

function handleFile(e) {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const file = files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const filas = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        procesarMatrizExcel(filas);
    };
    reader.readAsArrayBuffer(file);
}

function parseFechaTexto(texto) {
    // Handles "DD-MM-YYYY HH:MM" or "YYYY-MM-DD HH:MM"
    const partes = String(texto).trim().split(" ");
    const segFecha = partes[0];
    const segHora = partes[1] ? partes[1].substring(0, 5) : "00:00";
    const horaInt = parseInt(segHora.split(":")[0]) || 0;
    let fechaYMD = "";
    if (segFecha.includes("-")) {
        const p = segFecha.split("-");
        fechaYMD = p[0].length === 4
            ? `${p[0]}-${p[1]}-${p[2]}`
            : `${p[2]}-${p[1]}-${p[0]}`;
    } else if (segFecha.includes("/")) {
        const p = segFecha.split("/");
        fechaYMD = p[2].length === 4
            ? `${p[2]}-${p[1]}-${p[0]}`
            : `${p[0]}-${p[1]}-${p[2]}`;
    }
    return { fechaYMD, horaInt, horaCompletaStr: segHora };
}

function extraerFechaHora(celda) {
    if (!celda) return null;
    if (celda instanceof Date && !isNaN(celda)) {
        const tOffset = celda.getTimezoneOffset() * 60000;
        const d = new Date(celda.getTime() + tOffset);
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        const h = d.getHours();
        const min = String(d.getMinutes()).padStart(2, '0');
        return {
            fechaYMD: `${yyyy}-${mm}-${dd}`,
            horaInt: h,
            horaCompletaStr: `${String(h).padStart(2,'0')}:${min}`
        };
    }
    return parseFechaTexto(String(celda));
}

function esPendiente(val) {
    const v = String(val).trim().toUpperCase();
    // Duration pattern like "0 D 1 H 12 M" or blank/dash
    return !v || v === "-" || /^\d+ D \d+ H \d+ M$/.test(v);
}

function procesarMatrizExcel(filas) {
    datosGlobales = [];
    const fechasUnicas = new Set();
    const cuentaActual = document.getElementById('selectCuenta').value;

    // Columns: 0=ID, 1=Tipo, 2=Valor, 3=Criticidad, 4=Vehículo, 5=Viaje, 6=Tratada, 7=Evento, 8=Cerrado
    for (let i = 1; i < filas.length; i++) {
        const fila = filas[i];
        if (!fila || fila.length === 0) continue;

        const velocidadInt = parseFloat(String(fila[2] || "0").trim()) || 0;
        if (velocidadInt < 70) continue; // Solo alertas ≥70 km/h relevantes

        const tratadaRaw = fila[6] ? String(fila[6]).trim() : "";
        const pendiente = esPendiente(tratadaRaw);
        const colaborador = pendiente
            ? "SIN TRATAR / PENDIENTE"
            : tratadaRaw.toUpperCase();

        // Use col 7 (Evento) as the event timestamp — always populated
        const eventoInfo = extraerFechaHora(fila[7]);
        if (!eventoInfo || !eventoInfo.fechaYMD || eventoInfo.fechaYMD.length !== 10 || eventoInfo.fechaYMD.includes("undefined")) continue;

        const { fechaYMD, horaInt, horaCompletaStr } = eventoInfo;

        let turnoAsignado = "";
        if (horaInt >= 8 && horaInt < 16) turnoAsignado = "MAÑANA";
        else if (horaInt >= 16 && horaInt < 24) turnoAsignado = "TARDE";
        else turnoAsignado = "NOCHE";

        // For treated alerts: hora de gestión from col 8 (Cerrado)
        let horaGestion = "-";
        if (!pendiente && fila[8] && String(fila[8]).trim() !== "-") {
            const cerradoInfo = extraerFechaHora(fila[8]);
            if (cerradoInfo) horaGestion = cerradoInfo.horaCompletaStr;
            else horaGestion = horaCompletaStr;
        }

        fechasUnicas.add(fechaYMD);
        datosGlobales.push({
            colaborador,
            velocidad: velocidadInt,
            esTratado: !pendiente,
            fecha: fechaYMD,
            turno: turnoAsignado,
            cuenta: cuentaActual,
            horaGestion,
            horaEvento: horaCompletaStr,
            horaEventoInt: horaInt
        });
    }

    const selectF = document.getElementById('selectFecha');
    selectF.innerHTML = '<option value="">Seleccione un día...</option>';
    Array.from(fechasUnicas).sort().reverse().forEach(f => {
        const option = document.createElement('option');
        option.value = f;
        option.textContent = formatFechaVisual(f);
        selectF.appendChild(option);
    });

    if (datosGlobales.length > 0) {
        selectF.disabled = false;
        document.getElementById('selectTurno').disabled = true;
        const badge = document.getElementById('badgeCuentaActiva');
        badge.textContent = `CUENTA: ${cuentaActual}`;
        badge.className = `cuenta-badge fw-bold ${cuentaActual === 'DISTRIBUCION' ? 'bg-distribucion' : 'bg-proyectos'}`;
        document.getElementById('panelCuenta').style.display = 'block';
        alert(`Base de ${cuentaActual} sincronizada. ${datosGlobales.length} alertas (≥70 km/h) leídas.`);
    } else {
        alert("Verifique que el formato del reporte tenga los datos en las columnas correspondientes.");
    }
}

function formatFechaVisual(fechaStr) {
    const p = fechaStr.split("-");
    return `${p[2]}/${p[1]}/${p[0]}`;
}

function callbackFiltroFecha() {
    const fechaSel = document.getElementById('selectFecha').value;
    const selectT = document.getElementById('selectTurno');
    if (fechaSel) { selectT.disabled = false; selectT.value = ""; }
    else selectT.disabled = true;
    document.getElementById('statsRow').style.display = 'none';
    document.getElementById('dataRow').style.display = 'none';
}

function procesarDatosPantalla() {
    const fechaSeleccionada = document.getElementById('selectFecha').value;
    const turnoSeleccionado = document.getElementById('selectTurno').value;
    const cuentaActual = document.getElementById('selectCuenta').value;
    if (!fechaSeleccionada || !turnoSeleccionado) return;

    const datosFiltrados = datosGlobales.filter(d =>
        d.fecha === fechaSeleccionada &&
        d.turno === turnoSeleccionado &&
        d.cuenta === cuentaActual
    );

    let totalEventos = datosFiltrados.length;
    let totalTratados = 0, totalNoTratados = 0;
    const rendimiento = {}, ultimasHoras = {};
    let ultimaGestionMax = "";
    let tramo70=0, tramo80=0, tramo90=0;
    let pTr70=0, pTr80=0, pTr90=0;
    const pendientesPorHora = {};
    let velMaxPend = 0;

    datosFiltrados.forEach(d => {
        // Tramos generales (tratadas + pendientes)
        if (d.velocidad >= 70 && d.velocidad <= 79) tramo70++;
        else if (d.velocidad >= 80 && d.velocidad <= 89) tramo80++;
        else if (d.velocidad >= 90) tramo90++;

        if (d.esTratado) {
            totalTratados++;
            rendimiento[d.colaborador] = (rendimiento[d.colaborador] || 0) + 1;
            if (d.horaGestion !== "-") {
                if (!ultimasHoras[d.colaborador] || d.horaGestion > ultimasHoras[d.colaborador])
                    ultimasHoras[d.colaborador] = d.horaGestion;
                if (!ultimaGestionMax || d.horaGestion > ultimaGestionMax)
                    ultimaGestionMax = d.horaGestion;
            }
        } else {
            totalNoTratados++;
            // Pendientes por tramo
            if (d.velocidad >= 70 && d.velocidad <= 79) pTr70++;
            else if (d.velocidad >= 80 && d.velocidad <= 89) pTr80++;
            else if (d.velocidad >= 90) pTr90++;
            // Pendientes por hora
            const h = String(d.horaEventoInt).padStart(2,'0') + ":00";
            pendientesPorHora[h] = (pendientesPorHora[h] || 0) + 1;
            if (d.velocidad > velMaxPend) velMaxPend = d.velocidad;
        }
    });

    const listaOrdenada = Object.keys(rendimiento).map(key => ({
        nombre: key,
        cantidad: rendimiento[key],
        ultimaHora: ultimasHoras[key] || "Sin datos"
    })).sort((a, b) => b.cantidad - a.cantidad);

    // Hora pico de pendientes
    let horaPico = "-";
    if (Object.keys(pendientesPorHora).length > 0) {
        horaPico = Object.entries(pendientesPorHora).sort((a,b) => b[1]-a[1])[0][0];
    }

    // Update KPIs
    document.getElementById('txtTotalEventos').textContent = totalEventos;
    document.getElementById('txtTotalTratados').textContent = totalTratados;
    document.getElementById('txtTotalNoTratados').textContent = totalNoTratados;
    document.getElementById('txtTopColaborador').textContent = listaOrdenada.length > 0 ? listaOrdenada[0].nombre : "-";
    document.getElementById('lblMetaFecha').textContent = formatFechaVisual(fechaSeleccionada);
    document.getElementById('lblMetaTratada').textContent = ultimaGestionMax ? `${ultimaGestionMax} hrs` : "Sin gestión";

    // Update pendientes summary
    document.getElementById('badgePend70').textContent = pTr70;
    document.getElementById('badgePend80').textContent = pTr80;
    document.getElementById('badgePend90').textContent = pTr90;
    document.getElementById('lblHoraPico').textContent = horaPico !== "-" ? `${horaPico} hrs` : "-";
    document.getElementById('lblVelMaxPend').textContent = velMaxPend > 0 ? `${velMaxPend} km/h` : "-";

    actualizarTablaResumen(listaOrdenada);
    construirGraficos(listaOrdenada, tramo70, tramo80, tramo90, cuentaActual, pTr70, pTr80, pTr90, pendientesPorHora);
}

function actualizarTablaResumen(lista) {
    const tbody = document.getElementById('tablaResumen');
    tbody.innerHTML = "";
    if (lista.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="text-center text-muted">Sin alertas gestionadas</td></tr>`;
        return;
    }
    lista.forEach(colab => {
        const fila = document.createElement('tr');
        fila.innerHTML = `
            <td><strong>${colab.nombre}</strong></td>
            <td class="text-center"><span class="badge bg-secondary fs-6">${colab.cantidad}</span></td>
            <td class="text-center text-muted fw-bold">${colab.ultimaHora} hrs</td>
        `;
        tbody.appendChild(fila);
    });
}

function construirGraficos(lista, t70, t80, t90, cuenta, pTr70, pTr80, pTr90, pendHora) {
    const colorBarra = cuenta === 'DISTRIBUCION' ? 'rgba(253, 126, 20, 0.8)' : 'rgba(32, 201, 151, 0.8)';
    const colorBorde = cuenta === 'DISTRIBUCION' ? '#fd7e14' : '#20c997';

    // Chart 1: Barras tratadas por colaborador
    const ctxBarras = document.getElementById('chartProgreso').getContext('2d');
    if (miGraficoBarras) miGraficoBarras.destroy();
    miGraficoBarras = new Chart(ctxBarras, {
        type: 'bar',
        data: {
            labels: lista.map(c => c.nombre),
            datasets: [{ data: lista.map(c => c.cantidad), backgroundColor: colorBarra, borderColor: colorBorde, borderWidth: 1 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } },
            plugins: { legend: { display: false } }
        }
    });

    // Chart 2: Donut tramos (total)
    const ctxTramos = document.getElementById('chartTramos').getContext('2d');
    if (miGraficoTramos) miGraficoTramos.destroy();
    miGraficoTramos = new Chart(ctxTramos, {
        type: 'doughnut',
        data: {
            labels: ['Tramo 70-79 km/h', 'Tramo 80-89 km/h', 'Tramo ≥ 90 km/h'],
            datasets: [{
                data: [t70, t80, t90],
                backgroundColor: ['rgba(255,193,7,0.85)', 'rgba(253,126,20,0.85)', 'rgba(220,53,69,0.85)'],
                borderColor: ['#ffc107', '#fd7e14', '#dc3545'], borderWidth: 1
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } } }
        }
    });

    // Chart 3: Donut pendientes por tramo
    const ctxPendTramo = document.getElementById('chartPendientesTramo').getContext('2d');
    if (miGraficoPendTramo) miGraficoPendTramo.destroy();
    miGraficoPendTramo = new Chart(ctxPendTramo, {
        type: 'doughnut',
        data: {
            labels: ['Pendientes 70-79 km/h', 'Pendientes 80-89 km/h', 'Pendientes ≥ 90 km/h'],
            datasets: [{
                data: [pTr70, pTr80, pTr90],
                backgroundColor: ['rgba(255,193,7,0.9)', 'rgba(253,126,20,0.9)', 'rgba(220,53,69,0.9)'],
                borderColor: ['#ffc107', '#fd7e14', '#dc3545'], borderWidth: 2
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } },
                tooltip: {
                    callbacks: {
                        label: function(ctx) {
                            const total = pTr70 + pTr80 + pTr90;
                            const pct = total > 0 ? ((ctx.parsed / total)*100).toFixed(1) : 0;
                            return ` ${ctx.label}: ${ctx.parsed} (${pct}%)`;
                        }
                    }
                }
            }
        }
    });

    // Chart 4: Barras pendientes por hora
    const horasOrdenadas = Array.from({length:24}, (_,i) => String(i).padStart(2,'0')+":00");
    const valuesPorHora = horasOrdenadas.map(h => pendHora[h] || 0);
    const maxPend = Math.max(...valuesPorHora);

    const ctxPendHora = document.getElementById('chartPendientesHora').getContext('2d');
    if (miGraficoPendHora) miGraficoPendHora.destroy();
    miGraficoPendHora = new Chart(ctxPendHora, {
        type: 'bar',
        data: {
            labels: horasOrdenadas,
            datasets: [{
                label: 'Alertas pendientes',
                data: valuesPorHora,
                backgroundColor: valuesPorHora.map(v => v === maxPend && v > 0 ? 'rgba(220,53,69,0.9)' : 'rgba(220,53,69,0.45)'),
                borderColor: '#dc3545',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                x: { ticks: { maxRotation: 45, font: { size: 10 } } },
                y: { beginAtZero: true, ticks: { stepSize: 1 }, title: { display: true, text: 'Cant. alertas' } }
            },
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.y} alertas pendientes` } }
            }
        }
    });

    document.getElementById('statsRow').style.display = 'flex';
    document.getElementById('dataRow').style.display = 'block';
}
