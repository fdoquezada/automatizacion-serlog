// =========================================================================
// VARIABLES GLOBALES
// =========================================================================
let datosGlobales = [];
let datosTMS = {}; // Almacenar datos del TMS indexados por viaje
let miGraficoBarras = null;
let miGraficoTramos = null;
let miGraficoPendTramo = null;
let miGraficoPendHora = null;

// Variables globales para el control interactivo del modal de pendientes
let listaPendientesTemporal = []; 
let ordenActual = { columna: 'velocidad', ascendente: false };

// =========================================================================
// EVENT LISTENERS & INITIALIZATION
// =========================================================================
document.getElementById('excelFile').addEventListener('change', handleFile, false);
document.getElementById('excelTMS').addEventListener('change', handleFileTMS, false);
document.getElementById('selectFecha').addEventListener('change', callbackFiltroFecha, false);
document.getElementById('selectTurno').addEventListener('change', procesarDatosPantalla, false);
document.getElementById('btnProcesarDash').addEventListener('click', procesarDatosPantalla, false);
document.getElementById('btnLimpiarDatos').addEventListener('click', limpiarFiltros, false);

autoInit();

function autoInit() {
    const topCol = document.getElementById('txtTopColaborador');
    if (topCol) topCol.textContent = 'MOVIMIENTOS NOCTURNOS';
}

function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function limpiarFiltros() {
    [miGraficoBarras, miGraficoTramos, miGraficoPendTramo, miGraficoPendHora].forEach(grafico => {
        if (grafico) grafico.destroy();
    });
    miGraficoBarras = null;
    miGraficoTramos = null;
    miGraficoPendTramo = null;
    miGraficoPendHora = null;

    document.getElementById('excelFile').value = "";
    document.getElementById('excelTMS').value = "";
    datosTMS = {};
    const badge = document.getElementById('badgeTMS');
    if (badge) badge.style.display = 'none';
    const selectF = document.getElementById('selectFecha');
    selectF.innerHTML = '<option value="">Seleccione un día...</option>';
    selectF.disabled = true;
    const selectT = document.getElementById('selectTurno');
    selectT.value = "";
    selectT.disabled = true;
    document.getElementById('panelCuenta').style.display = 'none';
    document.getElementById('statsRow').style.display = 'none';
    document.getElementById('dataRow').style.display = 'none';
    document.getElementById('btnProcesarDash').disabled = true;
    setText('txtTotalEventos', '0');
    setText('txtTotalTratados', '0');
    setText('txtTotalNoTratados', '0');
    setText('txtTopColaborador', 'MOVIMIENTOS NOCTURNOS');
    setText('lblMetaFecha', '-');
    setText('lblMetaTratada', '-');
    setText('badgePendTotal', '0');
    setText('lblHoraPico', '-');
    setText('lblVelMaxPend', '-');
    document.getElementById('tablaResumen').innerHTML = '';
    datosGlobales = [];
    listaPendientesTemporal = [];
}

// =========================================================================
// PROCESAMIENTO DE ARCHIVOS (EXCEL PRINCIPAL Y TMS)
// =========================================================================
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

function handleFileTMS(e) {
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
        procesarDatosTMS(filas);
    };
    reader.readAsArrayBuffer(file);
}

function procesarDatosTMS(filas) {
    datosTMS = {};
    if (!filas || filas.length < 2) {
        alert("El archivo TMS está vacío o no tiene datos.");
        return;
    }

    const headers = filas[0].map(h => String(h || "").toLowerCase().trim());
    console.log("Total de headers encontrados:", headers.length);

    let idxViaje = -1;
    let idxPrioridad = -1;
    let idxEstadoViaje = -1;

    for (let i = 0; i < headers.length; i++) {
        const h = headers[i];
        if (idxViaje === -1 && (h === 'id viaje' || h === 'idviaje' || h === 'viaje')) {
            idxViaje = i;
            console.log(`✓ Asignado Id Viaje en columna ${i}: "${filas[0][i]}"`);
        }
        if (idxPrioridad === -1 && h.includes('prioridad')) {
            idxPrioridad = i;
            console.log(`✓ Encontrado Prioridad en columna ${i}: "${filas[0][i]}"`);
        }
        if (idxEstadoViaje === -1 && (h === 'estado viaje' || h.includes('estado viaje') || h === 'estado')) {
            idxEstadoViaje = i;
            console.log(`✓ Encontrado Estado Viaje en columna ${i}: "${filas[0][i]}"`);
        }
    }

    if (idxViaje === -1) {
        for (let i = 0; i < headers.length; i++) {
            const h = headers[i];
            if (h === 'id pedido' || h === 'idpedido' || h === 'pedido') {
                idxViaje = i;
                console.log(`⚠ Alerta: No se halló Id Viaje. Usando Id Pedido en columna ${i}: "${filas[0][i]}"`);
                break;
            }
        }
    }

    if (idxViaje === -1) {
        const headersTxt = filas[0].map((h, i) => `[${i}] ${h}`).join("\n");
        alert(`❌ No se pudo procesar el cruce.\n\nNo encontramos columnas de "Id Viaje" ni "Id Pedido".\n\nColumnas detectadas:\n${headersTxt}`);
        return;
    }

    let conteo = 0;
    for (let i = 1; i < filas.length; i++) {
        const fila = filas[i];
        if (!fila || fila.length === 0) continue;
        
        const viaje = String(fila[idxViaje] || "").trim();
        if (viaje) {
            datosTMS[viaje] = {
                prioridad: idxPrioridad !== -1 ? String(fila[idxPrioridad] || "-").trim() : "Normal",
                estadoViaje: idxEstadoViaje !== -1 ? String(fila[idxEstadoViaje] || "-").trim() : "Sin Estado"
            };
            conteo++;
        }
    }

    const badge = document.getElementById('badgeTMS');
    if (badge) {
        badge.textContent = `✓ Cargado (${conteo} viajes)`;
        badge.style.display = 'inline';
    }
    console.log("✅ Datos TMS vinculados correctamente:", conteo, "viajes indexados.");
}

// =========================================================================
// PARSERS DE FECHAS Y HORAS
// =========================================================================
function parseFechaTexto(texto) {
    texto = String(texto).trim();
    if (!texto) return { fechaYMD: '', horaInt: 0, horaCompletaStr: '00:00' };

    texto = texto.replace(/\s*hrs?\.?$/i, '').trim();
    texto = texto.replace(/\s*\b(AM|PM)\.?$/i, '').trim();

    const dateTimeMatch = texto.match(/^(\d{1,4}[\/\-]\d{1,2}[\/\-]\d{2,4})\s+(\d{1,2}:\d{2}(?::\d{2})?)$/);
    if (dateTimeMatch) {
        const fecha = dateTimeMatch[1];
        const hora = dateTimeMatch[2];
        const p = fecha.includes('-') ? fecha.split('-') : fecha.split('/');
        let fechaYMD = '';
        if (p.length === 3) {
            if (p[0].length === 4) fechaYMD = `${p[0]}-${String(p[1]).padStart(2, '0')}-${String(p[2]).padStart(2, '0')}`;
            else fechaYMD = `${p[2]}-${String(p[1]).padStart(2, '0')}-${String(p[0]).padStart(2, '0')}`;
        }
        const [horaNum, minutoNum] = hora.split(':').map(x => parseInt(x, 10) || 0);
        const horaReal = horaNum % 24;
        return {
            fechaYMD,
            horaInt: horaReal,
            horaCompletaStr: `${String(horaReal).padStart(2, '0')}:${String(minutoNum).padStart(2, '0')}`
        };
    }

    const dateOnlyMatch = texto.match(/^(\d{1,4}[\/\-]\d{1,2}[\/\-]\d{2,4})$/);
    if (dateOnlyMatch) {
        texto = `${dateOnlyMatch[1]} 00:00`;
    }

    const timeOnlyMatch = texto.match(/^(\d{1,4}):(\d{2})(?::(\d{2}))?$/);
    if (timeOnlyMatch) {
        let horaSegmento = timeOnlyMatch[1];
        const minuto = parseInt(timeOnlyMatch[2], 10) || 0;
        let hora = 0;
        if (horaSegmento.length <= 2) {
            hora = parseInt(horaSegmento, 10) || 0;
        } else if (horaSegmento.length === 3) {
            hora = parseInt(horaSegmento.slice(0, 1), 10) || 0;
            horaSegmento = `${horaSegmento.slice(0, 1)}:${horaSegmento.slice(1)}`;
        } else {
            hora = parseInt(horaSegmento.slice(0, horaSegmento.length - 2), 10) || 0;
        }
        const horaReal = hora % 24;
        return {
            fechaYMD: '',
            horaInt: horaReal,
            horaCompletaStr: `${String(horaReal).padStart(2, '0')}:${String(minuto).padStart(2, '0')}`
        };
    }

    const dateFallback = new Date(texto);
    if (!isNaN(dateFallback.getTime())) {
        const yyyy = dateFallback.getFullYear();
        const mm = String(dateFallback.getMonth() + 1).padStart(2, '0');
        const dd = String(dateFallback.getDate()).padStart(2, '0');
        const h = dateFallback.getHours();
        const min = String(dateFallback.getMinutes()).padStart(2, '0');
        return {
            fechaYMD: `${yyyy}-${mm}-${dd}`,
            horaInt: h,
            horaCompletaStr: `${String(h).padStart(2, '0')}:${min}`
        };
    }

    const partes = texto.split(' ');
    const segFecha = partes[0] || '';
    let segHora = partes[1] ? partes[1].substring(0, 8) : '00:00';

    if (segHora.includes(':')) {
        const [hora, minuto] = segHora.split(':');
        const horaNum = parseInt(hora, 10);
        const minNum = parseInt(minuto, 10) || 0;
        segHora = `${String(horaNum).padStart(2, '0')}:${String(minNum).padStart(2, '0')}`;
    } else {
        segHora = '00:00';
    }

    let fechaYMD = '';
    if (segFecha.includes('-')) {
        const p = segFecha.split('-');
        if (p[0].length === 4) fechaYMD = `${p[0]}-${String(p[1]).padStart(2, '0')}-${String(p[2]).padStart(2, '0')}`;
        else fechaYMD = `${p[2]}-${String(p[1]).padStart(2, '0')}-${String(p[0]).padStart(2, '0')}`;
    } else if (segFecha.includes('/')) {
        const p = segFecha.split('/');
        if (p[2] && p[2].length === 4) fechaYMD = `${p[2]}-${String(p[1]).padStart(2, '0')}-${String(p[0]).padStart(2, '0')}`;
        else fechaYMD = `${p[0]}-${String(p[1]).padStart(2, '0')}-${String(p[2]).padStart(2, '0')}`;
    }

    const horaInt = parseInt(segHora.split(':')[0], 10) || 0;
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
    if (typeof celda === 'number') {
        const d = new Date(Math.round((celda - 25569) * 86400 * 1000));
        if (!isNaN(d.getTime())) {
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
    }
    return parseFechaTexto(String(celda));
}

function buscarFechaHoraEnFila(fila) {
    for (let j = 0; j < fila.length; j++) {
        const valor = fila[j];
        if (valor === null || valor === undefined || valor === '') continue;
        if (typeof valor === 'string' && !/[\d]{1,2}[:\-\/]/.test(valor)) continue;
        const info = extraerFechaHora(valor);
        if (info && info.fechaYMD && info.fechaYMD.length === 10 && !info.fechaYMD.includes('undefined')) {
            return info;
        }
    }
    return null;
}

function esPendiente(val) {
    const v = String(val).trim().toUpperCase();
    if (!v || v === "-") return true;
    if (/^\d+\s*D\s*\d+\s*H\s*\d+\s*M$/.test(v)) return true;
    return false;
}

function procesarMatrizExcel(filas) {
    datosGlobales = [];
    const fechasUnicas = new Set();
    const cuentaActual = document.getElementById('selectCuenta').value;

    for (let i = 1; i < filas.length; i++) {
        const fila = filas[i];
        if (!fila || fila.length === 0) continue;

        const velocidadInt = parseFloat(String(fila[2] || "0").trim()) || 0;
        const tratadaRaw = fila[6] ? String(fila[6]).trim() : "";
        const pendiente = esPendiente(tratadaRaw);

        const colaborador = pendiente
            ? "SIN TRATAR / PENDIENTE"
            : tratadaRaw.replace(/\s+/g, ' ').trim().toUpperCase() || 'MOVIMIENTOS NOCTURNOS';

        let eventoInfo = extraerFechaHora(fila[7]);
        if (!eventoInfo || !eventoInfo.fechaYMD || eventoInfo.fechaYMD.length !== 10 || eventoInfo.fechaYMD.includes("undefined")) {
            eventoInfo = buscarFechaHoraEnFila(fila);
        }
        if (!eventoInfo || !eventoInfo.fechaYMD || eventoInfo.fechaYMD.length !== 10 || eventoInfo.fechaYMD.includes("undefined")) continue;

        const { fechaYMD, horaInt, horaCompletaStr } = eventoInfo;

        let turnoAsignado = "";
        if (horaInt >= 8 && horaInt < 16) turnoAsignado = "MAÑANA";
        else if (horaInt >= 16 && horaInt < 24) turnoAsignado = "TARDE";
        else turnoAsignado = "NOCHE";

        let horaGestion = "-";
        let turnoGestion = turnoAsignado;
        if (!pendiente) {
            if (fila[8] && String(fila[8]).trim() !== "-") {
                const cerradoInfo = extraerFechaHora(fila[8]);
                if (cerradoInfo && cerradoInfo.horaCompletaStr) {
                    horaGestion = cerradoInfo.horaCompletaStr;
                    const hc = cerradoInfo.horaInt;
                    if (hc >= 8 && hc < 16) turnoGestion = "MAÑANA";
                    else if (hc >= 16 && hc < 24) turnoGestion = "TARDE";
                    else turnoGestion = "NOCHE";
                } else {
                    horaGestion = horaCompletaStr;
                }
            } else {
                horaGestion = horaCompletaStr;
            }
        }

        fechasUnicas.add(fechaYMD);
        datosGlobales.push({
            colaborador,
            velocidad: velocidadInt,
            esTratado: !pendiente,
            fecha: fechaYMD,
            turno: turnoAsignado,
            turnoGestion,
            cuenta: cuentaActual,
            horaGestion,
            horaEvento: horaCompletaStr,
            horaEventoInt: horaInt,
            vehiculo: fila[4] ? String(fila[4]).trim() : "-",
            viaje: fila[5] ? String(fila[5]).trim() : "-"
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
        document.getElementById('btnProcesarDash').disabled = true;
        const badge = document.getElementById('badgeCuentaActiva');
        badge.textContent = `CUENTA: ${cuentaActual}`;
        badge.className = `cuenta-badge fw-bold ${cuentaActual === 'DISTRIBUCION' ? 'bg-distribucion' : 'bg-proyectos'}`;
        document.getElementById('panelCuenta').style.display = 'block';
        alert(`Base de ${cuentaActual} sincronizada. ${datosGlobales.length} alarmas leídas. Selecciona fecha y turno, luego presiona Procesar dashboard.`);
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
    const btnProcesar = document.getElementById('btnProcesarDash');
    if (fechaSel) {
        selectT.disabled = false;
        selectT.value = "";
        btnProcesar.disabled = false;
    } else {
        selectT.disabled = true;
        btnProcesar.disabled = true;
    }
    document.getElementById('statsRow').style.display = 'none';
    document.getElementById('dataRow').style.display = 'none';
}

// =========================================================================
// RENDERIZADO GENERAL DE LA PANTALLA PRINCIPAL
// =========================================================================
function procesarDatosPantalla() {
    const fechaSeleccionada = document.getElementById('selectFecha').value;
    const turnoSeleccionado = document.getElementById('selectTurno').value;
    const cuentaActual = document.getElementById('selectCuenta').value;
    if (!fechaSeleccionada) {
        alert('Selecciona primero un día.');
        return;
    }
    if (!turnoSeleccionado) {
        alert('Selecciona un turno antes de procesar.');
        return;
    }

    const datosFiltrados = datosGlobales.filter(d => {
        if (d.fecha !== fechaSeleccionada) return false;
        if (d.cuenta !== cuentaActual) return false;
        return d.turno === turnoSeleccionado;
    });

    if (datosFiltrados.length === 0) {
        alert('No hay datos para esta fecha/turno. Verifique la hora de evento y el turno seleccionado.');
        document.getElementById('statsRow').style.display = 'none';
        document.getElementById('dataRow').style.display = 'none';
        return;
    }

    let totalEventos = datosFiltrados.length;
    let totalTratados = 0, totalNoTratados = 0;
    const rendimiento = {}, ultimasHoras = {};
    let ultimaGestionMax = "";
    let tramoMenor70 = 0, tramo70=0, tramo80=0, tramo90=0;
    let pTrMenor70 = 0, pTr70=0, pTr80=0, pTr90=0;
    const pendientesPorHora = {};
    let velMaxPend = 0;

    datosFiltrados.forEach(d => {
        if (d.velocidad < 70) tramoMenor70++;
        else if (d.velocidad <= 79) tramo70++;
        else if (d.velocidad <= 89) tramo80++;
        else tramo90++;

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
            if (d.velocidad < 70) pTrMenor70++;
            else if (d.velocidad <= 79) pTr70++;
            else if (d.velocidad <= 89) pTr80++;
            else pTr90++;
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

    let horaPico = "-";
    if (Object.keys(pendientesPorHora).length > 0) {
        horaPico = Object.entries(pendientesPorHora).sort((a,b) => b[1]-a[1])[0][0];
    }

    setText('txtTotalEventos', totalEventos);
    setText('txtTotalTratados', totalTratados);
    setText('txtTotalNoTratados', totalNoTratados);
    setText('txtTopColaborador', listaOrdenada.length > 0 ? listaOrdenada[0].nombre : "-");
    setText('lblMetaFecha', formatFechaVisual(fechaSeleccionada));
    setText('lblMetaTratada', ultimaGestionMax ? `${ultimaGestionMax} hrs` : "Sin gestión");

    setText('badgePendTotal', totalNoTratados);
    setText('lblHoraPico', horaPico !== "-" ? `${horaPico} hrs` : "-");
    setText('lblVelMaxPend', velMaxPend > 0 ? `${velMaxPend} km/h` : "-");

    actualizarTablaResumen(listaOrdenada);
    construirGraficos(listaOrdenada, tramoMenor70, tramo70, tramo80, tramo90, cuentaActual, pTrMenor70, pTr70, pTr80, pTr90, pendientesPorHora);
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

// =========================================================================
// CONSTRUCCIÓN DE GRÁFICOS (CHART.JS)
// =========================================================================
function construirGraficos(lista, tMenor70, t70, t80, t90, cuenta, pMenor70, pTr70, pTr80, pTr90, pendHora) {
    const colorBarra = cuenta === 'DISTRIBUCION' ? 'rgba(253, 126, 20, 0.8)' : 'rgba(32, 201, 151, 0.8)';
    const colorBorde = cuenta === 'DISTRIBUCION' ? '#fd7e14' : '#20c997';

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

    const ctxTramos = document.getElementById('chartTramos').getContext('2d');
    if (miGraficoTramos) miGraficoTramos.destroy();
    miGraficoTramos = new Chart(ctxTramos, {
        type: 'doughnut',
        data: {
            labels: ['< 70 km/h', '70-79 km/h', '80-89 km/h', '≥ 90 km/h'],
            datasets: [{
                data: [tMenor70, t70, t80, t90],
                backgroundColor: ['rgba(40,167,69,0.85)', 'rgba(255,193,7,0.85)', 'rgba(253,126,20,0.85)', 'rgba(220,53,69,0.85)'],
                borderColor: ['#28a745', '#ffc107', '#fd7e14', '#dc3545'], borderWidth: 1
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } } }
        }
    });

    const ctxPendTramo = document.getElementById('chartPendientesTramo').getContext('2d');
    if (miGraficoPendTramo) miGraficoPendTramo.destroy();
    miGraficoPendTramo = new Chart(ctxPendTramo, {
        type: 'doughnut',
        data: {
            labels: ['< 70 km/h', '70-79 km/h', '80-89 km/h', '≥ 90 km/h'],
            datasets: [{
                data: [pMenor70, pTr70, pTr80, pTr90],
                backgroundColor: ['rgba(40,167,69,0.9)', 'rgba(255,193,7,0.9)', 'rgba(253,126,20,0.9)', 'rgba(220,53,69,0.9)'],
                borderColor: ['#28a745', '#ffc107', '#fd7e14', '#dc3545'], borderWidth: 2
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } },
                tooltip: {
                    callbacks: {
                        label: function(ctx) {
                            const total = pMenor70 + pTr70 + pTr80 + pTr90;
                            const pct = total > 0 ? ((ctx.parsed / total)*100).toFixed(1) : 0;
                            return ` ${ctx.label}: ${ctx.parsed} (${pct}%)`;
                        }
                    }
                }
            }
        }
    });

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

    const btnVerPendientes = document.getElementById('btnVerPendientes');
    if (btnVerPendientes) {
        btnVerPendientes.style.cursor = 'pointer';
        btnVerPendientes.onclick = mostrarDetallePendientes;
    }
}

// =========================================================================
// =========================================================================
// CONTROLADOR INTERACTIVO DEL MODAL DE DETALLES PENDIENTES
// =========================================================================
function mostrarDetallePendientes() {
    const fechaSeleccionada = document.getElementById('selectFecha').value;
    const turnoSeleccionado = document.getElementById('selectTurno').value;
    const cuentaActual = document.getElementById('selectCuenta').value;

    // Guardar el filtro inicial en la lista temporal compartida
    listaPendientesTemporal = datosGlobales.filter(d => {
        if (d.fecha !== fechaSeleccionada || d.cuenta !== cuentaActual) return false;
        return !d.esTratado && d.turno === turnoSeleccionado;
    });

    // Criterio de orden por defecto: Ahora iniciará ordenando por ID Pedido de forma ascendente
    listaPendientesTemporal.sort((a, b) => {
        const idA = String(a.viaje || '').split('|')[0].trim();
        const idB = String(b.viaje || '').split('|')[0].trim();
        return idA.localeCompare(idB, undefined, { numeric: true, sensitivity: 'base' });
    });
    ordenActual = { columna: 'viaje', ascendente: true };

    const modal = document.getElementById('modalPendientes');
    const titulo = document.getElementById('modalLabel');
    titulo.textContent = 'Alertas Pendientes - Listado';

    // Renderizar la tabla interactiva por primera vez
    renderizarTablaModal();

    // Sincronizar datos por si el usuario presiona descargar inmediatamente
    window.pendientesActuales = {
        rango: 'Pendientes totales',
        fecha: formatFechaVisual(fechaSeleccionada),
        turno: turnoSeleccionado,
        datos: listaPendientesTemporal
    };

    const bsModal = new bootstrap.Modal(modal);
    bsModal.show();

    const btnExportar = document.getElementById('btnExportarPendientes');
    if (btnExportar) {
        btnExportar.onclick = exportarPendientesExcel;
    }
}

// =========================================================================
// MOTOR DE RENDIMIENTO DEL MODAL (CON SOPORTE DE ID PEDIDO REEMPLAZADO)
// =========================================================================
function renderizarTablaModal() {
    const content = document.getElementById('modalPendientesContent');
    if (!listaPendientesTemporal || listaPendientesTemporal.length === 0) {
        content.innerHTML = `<div class="alert alert-success text-center py-4"><strong>✓ Sin alertas pendientes</strong> para este turno.</div>`;
        return;
    }

    // Dibujar estados dinámicos en las cabeceras según la interacción actual
    const flecha = (col) => {
        if (ordenActual.columna !== col) return ' <span style="opacity: 0.4;">↕</span>';
        return ordenActual.ascendente ? ' ▲' : ' ▼';
    };

    let html = `
        <div class="table-responsive">
            <table class="table table-striped table-hover align-middle table-sm">
                <thead class="table-danger" style="user-select: none;">
                    <tr>
                        <th style="cursor:pointer;" onclick="ordenarColumnaModal('indice')">#${flecha('indice')}</th>
                        <th style="cursor:pointer;" onclick="ordenarColumnaModal('viaje')">ID Pedido${flecha('viaje')}</th>
                        <th style="cursor:pointer;" onclick="ordenarColumnaModal('vehiculo')">Vehículo${flecha('vehiculo')}</th>
                        <th style="cursor:pointer;" onclick="ordenarColumnaModal('horaEvento')">Hora Evento${flecha('horaEvento')}</th>
                        <th style="cursor:pointer;" onclick="ordenarColumnaModal('prioridad')">Prioridad${flecha('prioridad')}</th>
                        <th style="cursor:pointer;" onclick="ordenarColumnaModal('estadoViaje')">Estado Viaje${flecha('estadoViaje')}</th>
                        <th>Estado</th>
                    </tr>
                </thead>
                <tbody>
    `;

    listaPendientesTemporal.forEach((d, idx) => {
        const identificadorOriginal = d.viaje || d['Id Viaje'] || d['Id Pedido'] || "-";
        // Extraemos limpiamente el número largo (ej. 139849132)
        const viajeNumero = String(identificadorOriginal).split('|')[0].trim();
        
        const datosViajesTMS = datosTMS[viajeNumero] || {};
        const prioridad = datosViajesTMS.prioridad || '-';
        const estadoViaje = datosViajesTMS.estadoViaje || '-';

        const badgePrioridad = prioridad === 'Urgencia' ? 'bg-info text-white' : 'bg-secondary text-white';
        const badgeEstadoViaje = estadoViaje === 'Activo' ? 'bg-warning text-dark' : 'bg-light text-dark';
        
        html += `
            <tr>
                <td><strong>${idx + 1}</strong></td>
                <td><code class="fs-6" style="color:#d63384; font-weight: bold; user-select: all;">${viajeNumero}</code></td>
                <td><code style="font-size:0.75rem; color:#6c757d;">${d.vehiculo}</code></td>
                <td>${d.horaEvento}</td>
                <td><span class="badge ${badgePrioridad}">${prioridad}</span></td>
                <td><span class="badge ${badgeEstadoViaje}">${estadoViaje}</span></td>
                <td><span class="badge bg-secondary">Pendiente</span></td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
        </div>
        <div class="alert alert-info mt-3 mb-0">
            <i class="bi bi-info-circle"></i>
            <strong>${listaPendientesTemporal.length}</strong> alertas sin tratar para este turno.
        </div>
    `;
    content.innerHTML = html;
}

// =========================================================================
// ORDENAMIENTO EN TIEMPO REAL DESDE CABECERAS
// =========================================================================
function ordenarColumnaModal(columna) {
    if (ordenActual.columna === columna) {
        ordenActual.ascendente = !ordenActual.ascendente;
    } else {
        ordenActual.columna = columna;
        ordenActual.ascendente = true;
    }

    listaPendientesTemporal.sort((a, b) => {
        let valA, valB;

        if (columna === 'prioridad') {
            const viajeA = String(a.viaje || '').split('|')[0].trim();
            const viajeB = String(b.viaje || '').split('|')[0].trim();
            valA = datosTMS[viajeA]?.prioridad || '-';
            valB = datosTMS[viajeB]?.prioridad || '-';
        } else if (columna === 'estadoViaje') {
            const viajeA = String(a.viaje || '').split('|')[0].trim();
            const viajeB = String(b.viaje || '').split('|')[0].trim();
            valA = datosTMS[viajeA]?.estadoViaje || '-';
            valB = datosTMS[viajeB]?.estadoViaje || '-';
        } else if (columna === 'indice') {
            return ordenActual.ascendente ? 1 : -1;
        } else if (columna === 'viaje') {
            // Ordenar de forma natural por el ID numérico largo
            valA = String(a.viaje || '').split('|')[0].trim();
            valB = String(b.viaje || '').split('|')[0].trim();
            return ordenActual.ascendente 
                ? valA.localeCompare(valB, undefined, { numeric: true }) 
                : valB.localeCompare(valA, undefined, { numeric: true });
        } else {
            valA = String(a[columna] || '').toUpperCase();
            valB = String(b[columna] || '').toUpperCase();
        }

        if (valA < valB) return ordenActual.ascendente ? -1 : 1;
        if (valA > valB) return ordenActual.ascendente ? 1 : -1;
        return 0;
    });

    if (window.pendientesActuales) {
        window.pendientesActuales.datos = listaPendientesTemporal;
    }

    renderizarTablaModal();
}

// =========================================================================
// EXPORTACIÓN EXCEL DE LOS DATOS REORDENADOS
// =========================================================================
function exportarPendientesExcel() {
    if (!window.pendientesActuales || !window.pendientesActuales.datos) {
        alert("No hay datos para exportar");
        return;
    }

    const data = window.pendientesActuales;
    const ws_data = [
        ['REPORTE DE ALERTAS PENDIENTES'],
        [],
        ['Fecha:', data.fecha],
        ['Turno:', data.turno],
        ['Detalle:', data.rango],
        ['Total Pendientes:', data.datos.length],
        [],
        ['#', 'ID Pedido', 'Vehículo', 'Hora Evento', 'Estado', 'Prioridad', 'Estado Viaje']
    ];

    data.datos.forEach((d, idx) => {
        const identificadorOriginal = d.viaje || d['Id Viaje'] || d['Id Pedido'] || "-";
        const viajeNumero = String(identificadorOriginal).split('|')[0].trim();
        
        const datosViajesTMS = datosTMS[viajeNumero] || {};
        
        ws_data.push([
            idx + 1,
            viajeNumero,
            d.vehiculo,
            d.horaEvento,
            'Pendiente',
            datosViajesTMS.prioridad || '-',
            datosViajesTMS.estadoViaje || '-'
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    ws['!cols'] = [
        { wch: 8 }, 
        { wch: 20 }, 
        { wch: 20 }, 
        { wch: 15 },
        { wch: 15 },
        { wch: 15 },
        { wch: 18 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pendientes");
    const fileName = `Pendientes_${data.fecha}_${data.turno}_Resumen.xlsx`;
    XLSX.writeFile(wb, fileName);
}