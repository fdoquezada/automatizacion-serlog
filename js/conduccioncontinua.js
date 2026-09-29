let registros = [], chartEstado, chartHoras, chartVehiculos, chartUsuarios, vistaActual = [];
const $ = id => document.getElementById(id);

['fechaDesde', 'fechaHasta'].forEach(id => $(id).addEventListener('change', prepararFiltros));
$('selectTurno').addEventListener('change', renderizar);
$('selectCuenta').addEventListener('change', limpiar);
$('excelFile').addEventListener('change', leerExcel);
$('btnExportar').addEventListener('click', exportarExcel);

function limpiar() {
    registros = [];
    vistaActual = [];
    ['fechaDesde', 'fechaHasta'].forEach(id => { $(id).value = ''; $(id).disabled = true; });
    $('selectTurno').value = '';
    $('selectTurno').disabled = true;
    $('statsRow').hidden = true;
    $('dataRow').hidden = true;
    mostrarInfo('', false);
}

function leerExcel(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
            const filas = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' });
            registros = filas.slice(1).map(normalizar).filter(Boolean);
            if (!registros.length) throw new Error('No se encontraron registros válidos.');
            const fechas = registros.map(r => r.fecha).sort();
            $('fechaDesde').value = fechas[0];
            $('fechaHasta').value = fechas[fechas.length - 1];
            $('fechaDesde').disabled = false;
            $('fechaHasta').disabled = false;
            prepararFiltros();
            mostrarInfo(`Excel cargado: ${registros.length.toLocaleString('es-CL')} alertas únicas.`, true);
        } catch (error) {
            mostrarInfo(error.message, true, 'danger');
        }
    };
    reader.readAsArrayBuffer(file);
}

function normalizar(row) {
    const evento = parseFecha(row[7]);
    if (!evento) return null;
    const tratada = String(row[6] ?? '').trim();
    const pendiente = !tratada || /^\d+\s*d\s*\d+\s*h\s*\d+\s*m$/i.test(tratada) || tratada === '-';
    return {
        id: String(row[0] ?? '').trim(),
        vehiculo: String(row[4] ?? '-').trim() || '-',
        viaje: String(row[5] ?? '-').trim() || '-',
        fecha: evento.fecha,
        hora: evento.hora,
        horaInt: evento.horaInt,
        evento: `${evento.fecha} ${evento.hora}`,
        tratado: !pendiente,
        gestor: pendiente ? '-' : tratada.replace(/\s+/g, ' ').trim().toUpperCase(),
        duracion: pendiente ? tratada : '-',
        minutos: pendiente ? duracionMinutos(tratada) : 0
    };
}

function parseFecha(value) {
    if (value instanceof Date && !isNaN(value)) {
        const date = new Date(value.getTime() - value.getTimezoneOffset() * 60000);
        return fechaObj(date.getFullYear(), date.getMonth() + 1, date.getDate(), date.getHours(), date.getMinutes());
    }
    const match = String(value ?? '').trim().match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})\s*(\d{1,2})?:?(\d{2})?/);
    return match ? fechaObj(match[3], match[2], match[1], match[4] || 0, match[5] || 0) : null;
}

function fechaObj(year, month, day, hour, minute) {
    return { fecha: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`, hora: `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`, horaInt: Number(hour) };
}

function duracionMinutos(value) {
    const match = String(value).match(/(\d+)\s*d\s*(\d+)\s*h\s*(\d+)\s*m/i);
    return match ? Number(match[1]) * 1440 + Number(match[2]) * 60 + Number(match[3]) : 0;
}

function prepararFiltros() {
    const desde = $('fechaDesde').value;
    const hasta = $('fechaHasta').value;
    $('selectTurno').disabled = !(desde && hasta && desde <= hasta);
    if ($('selectTurno').disabled) {
        $('statsRow').hidden = true;
        $('dataRow').hidden = true;
    } else if ($('selectTurno').value) renderizar();
}

function turno(hora) { return hora >= 8 && hora < 16 ? 'MAÑANA' : hora >= 16 ? 'TARDE' : 'NOCHE'; }

function renderizar() {
    const desde = $('fechaDesde').value;
    const hasta = $('fechaHasta').value;
    const turnoSeleccionado = $('selectTurno').value;
    if (!desde || !hasta || !turnoSeleccionado) return;
    vistaActual = registros.filter(r => r.fecha >= desde && r.fecha <= hasta && turno(r.horaInt) === turnoSeleccionado);
    const pendientes = vistaActual.filter(r => !r.tratado);
    $('txtTotalEventos').textContent = vistaActual.length;
    $('txtTotalTratados').textContent = vistaActual.length - pendientes.length;
    $('txtTotalPendientes').textContent = pendientes.length;
    const max = pendientes.reduce((total, row) => Math.max(total, row.minutos), 0);
    $('txtMaxDuracion').textContent = max ? formatoDuracion(max) : '-';
    $('statsRow').hidden = false;
    $('dataRow').hidden = false;
    construirGraficos(pendientes);
    construirPendientes(pendientes);
    construirUsuarios();
    construirTabla(vistaActual);
}

function construirGraficos(pendientes) {
    const horas = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));
    const porHora = horas.map(hora => vistaActual.filter(row => String(row.horaInt).padStart(2, '0') === hora).length);
    const conteo = {};
    vistaActual.forEach(row => { conteo[row.vehiculo] = (conteo[row.vehiculo] || 0) + 1; });
    const top = Object.entries(conteo).sort((a, b) => b[1] - a[1]).slice(0, 10);
    if (chartEstado) chartEstado.destroy();
    if (chartHoras) chartHoras.destroy();
    if (chartVehiculos) chartVehiculos.destroy();
    chartEstado = new Chart($('chartEstado'), { type: 'doughnut', data: { labels: ['Tratadas', 'Pendientes'], datasets: [{ data: [vistaActual.length - pendientes.length, pendientes.length], backgroundColor: ['#17865b', '#d93025'] }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } } });
    chartHoras = new Chart($('chartHoras'), { type: 'bar', data: { labels: horas.map(hora => `${hora}:00`), datasets: [{ label: 'Alertas', data: porHora, backgroundColor: '#d93025' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false } } } });
    chartVehiculos = new Chart($('chartVehiculos'), { type: 'bar', data: { labels: top.map(item => item[0]), datasets: [{ label: 'Alertas', data: top.map(item => item[1]), backgroundColor: '#c87900' }] }, options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, scales: { x: { beginAtZero: true, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false } } } });
}

function construirUsuarios() {
    const conteo = {};
    vistaActual.filter(row => row.tratado).forEach(row => { conteo[row.gestor] = (conteo[row.gestor] || 0) + 1; });
    const usuarios = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
    if (chartUsuarios) chartUsuarios.destroy();
    chartUsuarios = new Chart($('chartUsuarios'), { type: 'bar', data: { labels: usuarios.map(item => item[0]), datasets: [{ label: 'Gestiones', data: usuarios.map(item => item[1]), backgroundColor: usuarios.map((item, index) => index === usuarios.length - 1 ? '#d93025' : '#17865b') }] }, options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, scales: { x: { beginAtZero: true, ticks: { stepSize: 1 } } }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: context => ` ${context.parsed.x} gestiones realizadas` } } } } });
    const menores = usuarios.slice().sort((a, b) => a[1] - b[1]).slice(0, 6);
    $('userSummary').innerHTML = menores.length ? menores.map(([usuario, cantidad], index) => `<div class="user-line"><div><strong>${index + 1}. ${esc(usuario)}</strong><small>Gestiones realizadas</small></div><span class="user-badge ${index === 0 ? 'attention' : ''}">${cantidad}</span></div>`).join('') : '<p class="text-muted mb-0">No hay usuarios con gestiones en el filtro seleccionado.</p>';
}

function construirPendientes(pendientes) {
    const porVehiculo = {};
    pendientes.forEach(row => { porVehiculo[row.vehiculo] = (porVehiculo[row.vehiculo] || 0) + 1; });
    const top = Object.entries(porVehiculo).sort((a, b) => b[1] - a[1]).slice(0, 6);
    $('pendingSummary').innerHTML = top.length ? top.map(([vehiculo, cantidad]) => `<div class="pending-line"><div><strong>${esc(vehiculo)}</strong><small>Vehículo con alerta pendiente</small></div><span class="pending-badge">${cantidad}</span></div>`).join('') : '<p class="text-muted mb-0">No hay alertas pendientes en el filtro seleccionado.</p>';
}

function construirTabla(datos) {
    const filas = datos.slice().sort((a, b) => b.evento.localeCompare(a.evento)).slice(0, 200);
    $('tablaDetalle').innerHTML = filas.length ? filas.map(row => `<tr><td><code>${esc(row.id)}</code></td><td><strong>${esc(row.vehiculo)}</strong></td><td>${esc(row.viaje)}</td><td>${row.evento}</td><td>${esc(row.duracion)}</td><td><span class="status ${row.tratado ? 'done' : 'pending'}">${row.tratado ? 'Tratada' : 'Pendiente'}</span></td><td>${esc(row.gestor)}</td></tr>`).join('') : '<tr><td colspan="7" class="text-center text-muted py-4">Sin registros para el filtro seleccionado.</td></tr>';
}

function exportarExcel() {
    if (!vistaActual.length) return;
    const filas = vistaActual.map(row => ({ ID: row.id, Vehículo: row.vehiculo, Viaje: row.viaje, Evento: row.evento, Duración: row.duracion, Estado: row.tratado ? 'Tratada' : 'Pendiente', Gestor: row.gestor }));
    const hoja = XLSX.utils.json_to_sheet(filas);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Conduccion Continua');
    XLSX.writeFile(libro, `Conduccion_Continua_${$('fechaDesde').value}_${$('fechaHasta').value}.xlsx`);
}

function formatoDuracion(minutos) { const dias = Math.floor(minutos / 1440); const horas = Math.floor((minutos % 1440) / 60); const minutosRestantes = minutos % 60; return `${dias ? `${dias}d ` : ''}${horas}h ${minutosRestantes}m`; }
function esc(value) { return String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char])); }
function mostrarInfo(texto, visible, tipo = 'success') { const info = $('info'); info.textContent = texto; info.className = `alert dashboard-info ${visible ? 'show' : ''} alert-${tipo}`; }
