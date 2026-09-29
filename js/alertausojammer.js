(function () {
    'use strict';

    const PAGE_SIZE = 25;
    let datosGlobal = [];
    let datosFiltrados = [];
    let currentPage = 1;
    let charts = {};

    const text = value => value == null ? '' : String(value).trim();
    const normalize = value => text(value).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    const headerIndex = (headers, keys) => headers.findIndex(header => keys.some(key => header.includes(key)));

    function parseDate(value) {
        if (value instanceof Date && !Number.isNaN(value.getTime())) return formatDate(value);
        const valueText = text(value);
        const match = valueText.match(/(\d{1,2})[-/]([01]?\d)[-/](\d{2,4})\s+(\d{1,2}):(\d{2})/);
        if (!match) return { date: '-', time: valueText || '00:00', hour: 0 };
        const year = match[3].length === 2 ? `20${match[3]}` : match[3];
        return { date: `${year}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`, time: `${match[4].padStart(2, '0')}:${match[5]}`, hour: Number(match[4]) };
    }

    function formatDate(date) {
        return { date: `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`, time: `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`, hour: date.getHours() };
    }

    function getShift(hour, time) {
        const minutes = hour * 60 + Number((time || '00:00').split(':')[1] || 0);
        if (minutes >= 480 && minutes <= 960) return 'Mañana';
        if (minutes >= 961 && minutes <= 1439) return 'Tarde';
        return 'Noche';
    }

    function isTreated(value) {
        const result = text(value).toUpperCase();
        return Boolean(result && result !== '-' && !/^\d+\s*D\s*\d+\s*H\s*\d+\s*M$/.test(result) && result !== '0');
    }

    function processRows(rows) {
        if (!rows.length) return [];
        const headerRowIndex = rows.findIndex(row => {
            const headers = row.map(normalize);
            return headers.some(header => header.includes('evento') || header.includes('fecha')) && headers.some(header => header === 'tipo' || header.includes('alarma') || header === 'id');
        });
        if (headerRowIndex < 0) return [];
        const headers = rows[headerRowIndex].map(normalize);
        const indexes = {
            id: headerIndex(headers, ['id', 'codigo']), type: headerIndex(headers, ['tipo', 'alarma', 'alerta']), value: headerIndex(headers, ['valor']),
            criticality: headerIndex(headers, ['critic', 'prioridad']), vehicle: headerIndex(headers, ['veh', 'patente', 'matricula']), trip: headerIndex(headers, ['viaj', 'servicio']),
            treated: headerIndex(headers, ['trat', 'gestion', 'responsable']), event: headerIndex(headers, ['evento', 'fecha']), closed: headerIndex(headers, ['cerr', 'cierre'])
        };
        return rows.slice(headerRowIndex + 1).map(row => {
            const type = text(row[indexes.type]);
            if (!/jammer/i.test(type) && !/jammer/i.test(row.map(text).join(' '))) return null;
            const event = parseDate(row[indexes.event]);
            const closed = parseDate(row[indexes.closed]);
            const treated = isTreated(row[indexes.treated]);
            const responsible = treated && text(row[indexes.treated]).match(/[a-záéíóúñ]/i) ? text(row[indexes.treated]).toUpperCase() : '';
            return {
                id: text(row[indexes.id]), vehicle: text(row[indexes.vehicle]), trip: text(row[indexes.trip]),
                value: text(row[indexes.value]) || '-', criticality: text(row[indexes.criticality]) || 'Sin informar',
                eventDate: event.date, eventTime: event.time, eventHour: event.hour, closedDate: closed.date,
                treated, status: treated ? 'Tratada' : 'Pendiente', responsible: responsible || (treated ? 'Sin nombre' : 'Sin tratar'),
                eventShift: getShift(event.hour, event.time), treatmentShift: treated ? getShift(closed.hour, closed.time) : ''
            };
        }).filter(Boolean);
    }

    function selectedData() {
        const date = document.getElementById('selectFecha').value;
        const shift = document.getElementById('selectTurno').value;
        const status = document.getElementById('selectFiltroEstado').value;
        return datosGlobal.filter(item => (!date || item.eventDate === date) && (!shift || (item.treated ? item.treatmentShift : item.eventShift) === shift) && (status === 'todos' || (status === 'tratadas' ? item.treated : !item.treated)));
    }

    function render() {
        const date = document.getElementById('selectFecha').value;
        datosFiltrados = date ? selectedData() : [];
        currentPage = 1;
        const treated = datosFiltrados.filter(item => item.treated).length;
        const pending = datosFiltrados.length - treated;
        document.getElementById('txtTotalAlertas').textContent = datosFiltrados.length;
        document.getElementById('txtTotalTratadas').textContent = treated;
        document.getElementById('txtTotalPendientes').textContent = pending;
        updateCharts(treated, pending);
        renderTable();
        renderRanking();
        renderShiftSummary();
    }

    function renderTable() {
        const tbody = document.getElementById('tablaAlertasBody');
        const pageItems = datosFiltrados.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);
        tbody.innerHTML = pageItems.length ? pageItems.map(item => `<tr><td>${text(item.id)}</td><td>${text(item.vehicle)}</td><td>${text(item.trip)}</td><td class="criticality">${text(item.criticality)}</td><td>${text(item.eventDate)} ${text(item.eventTime)}</td><td><span class="badge ${item.treated ? 'bg-success' : 'bg-danger'}">${item.status}</span></td><td>${text(item.responsible)}</td></tr>`).join('') : '<tr><td colspan="7" class="text-center text-muted">No hay datos para mostrar</td></tr>';
        const pages = Math.max(1, Math.ceil(datosFiltrados.length / PAGE_SIZE));
        document.getElementById('paginationControls').innerHTML = `<div class="btn-group btn-group-sm"><button class="btn btn-outline-secondary" ${currentPage === 1 ? 'disabled' : ''} data-page="${currentPage - 1}">«</button>${Array.from({ length: pages }, (_, index) => index + 1).slice(Math.max(0, currentPage - 4), currentPage + 3).map(page => `<button class="btn ${page === currentPage ? 'btn-primary' : 'btn-outline-secondary'}" data-page="${page}">${page}</button>`).join('')}<button class="btn btn-outline-secondary" ${currentPage === pages ? 'disabled' : ''} data-page="${currentPage + 1}">»</button></div><span class="text-muted small">${datosFiltrados.length} registros</span>`;
        document.querySelectorAll('#paginationControls [data-page]').forEach(button => button.addEventListener('click', () => { currentPage = Number(button.dataset.page); renderTable(); }));
    }

    function renderRanking() {
        const ranking = {};
        datosFiltrados.filter(item => item.treated).forEach(item => { ranking[item.responsible] = (ranking[item.responsible] || 0) + 1; });
        const entries = Object.entries(ranking).sort((a, b) => b[1] - a[1]).slice(0, 6);
        document.getElementById('rankingTratadas').innerHTML = entries.length ? entries.map(([name, count], index) => `<div class="ranking-item"><strong>#${index + 1} ${text(name)}</strong><span class="badge bg-success">${count}</span></div>`).join('') : '<div class="text-muted small">No hay registros tratados.</div>';
        charts.ranking.data.labels = entries.map(entry => entry[0]);
        charts.ranking.data.datasets[0].data = entries.map(entry => entry[1]);
        charts.ranking.update();
    }

    function renderShiftSummary() {
        const date = document.getElementById('selectFecha').value;
        const shift = document.getElementById('selectTurno').value;
        const shiftData = datosGlobal.filter(item => item.eventDate === date && (item.treated ? item.treatmentShift : item.eventShift) === shift);
        const pending = shiftData.filter(item => !item.treated).length;
        const treated = shiftData.filter(item => item.treated).length;
        document.getElementById('txtPendientesTurno').textContent = shift ? pending : 0;
        document.getElementById('txtGestionesTurno').textContent = shift ? treated : 0;
        charts.shift.data.datasets[0].data = [pending, treated];
        charts.shift.update();
    }

    function updateCharts(treated, pending) {
        charts.total.data.datasets[0].data = [treated + pending, Math.max(1, treated + pending)];
        charts.total.update();
        charts.status.data.datasets[0].data = [treated, pending];
        charts.status.update();
    }

    function initCharts() {
        const doughnut = (id, colors, labels) => new Chart(document.getElementById(id), { type: 'doughnut', data: { labels, datasets: [{ data: [0, 0], backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { font: { size: 9 }, boxWidth: 12 } } } } });
        charts.total = new Chart(document.getElementById('chartTotalAlertas'), { type: 'doughnut', data: { labels: ['Alertas'], datasets: [{ data: [0, 1], backgroundColor: ['#0d6efd', '#e9ecef'], borderWidth: 0, cutout: '75%' }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } } });
        charts.status = doughnut('chartEstadoResumen', ['#198754', '#dc3545'], ['Tratadas', 'Pendientes']);
        charts.shift = doughnut('chartTurnoResumen', ['#dc3545', '#198754'], ['Pendientes', 'Gestionadas']);
        charts.ranking = new Chart(document.getElementById('chartRanking'), { type: 'bar', data: { labels: [], datasets: [{ label: 'Tratadas', data: [], backgroundColor: '#0d6efd', borderRadius: 4 }] }, options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true }, y: { grid: { display: false } } } } });
    }

    function populateFilters() {
        const dates = [...new Set(datosGlobal.map(item => item.eventDate).filter(date => date !== '-'))].sort().reverse();
        const dateSelect = document.getElementById('selectFecha');
        dateSelect.innerHTML = '<option value="">Seleccione un día</option>' + dates.map(date => `<option value="${date}">${date.split('-').reverse().join('/')}</option>`).join('');
        dateSelect.disabled = false;
        document.getElementById('selectFiltroEstado').disabled = false;
        document.getElementById('tablaAlertasBody').innerHTML = '<tr><td colspan="7" class="text-center text-muted">Seleccione un día para ver los datos.</td></tr>';
        updateShifts();
    }

    function updateShifts() {
        const date = document.getElementById('selectFecha').value;
        const shiftSelect = document.getElementById('selectTurno');
        const shifts = ['Mañana', 'Tarde', 'Noche'];
        shiftSelect.innerHTML = '<option value="">Seleccione un turno</option>' + shifts.map(shift => { const count = datosGlobal.filter(item => (!date || item.eventDate === date) && (item.eventShift === shift || item.treatmentShift === shift)).length; return `<option value="${shift}">${shift} (${count})</option>`; }).join('');
        shiftSelect.disabled = !date;
    }

    function loadWorkbook(arrayBuffer, source) {
        try {
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '' });
            datosGlobal = processRows(rows);
            populateFilters();
            document.getElementById('sourceStatus').textContent = `${source}: ${datosGlobal.length} alertas Jammer encontradas`;
            render();
        } catch (error) {
            document.getElementById('sourceStatus').textContent = 'No fue posible leer el libro Excel.';
            console.error(error);
        }
    }

    function exportExcel(items, filename) {
        if (!items.length) return alert('No hay datos para exportar.');
        const rows = items.map(item => [item.id, item.vehicle, item.trip, item.criticality, item.eventDate, item.eventTime, item.status, item.responsible, item.treated ? item.treatmentShift : item.eventShift]);
        const worksheet = XLSX.utils.aoa_to_sheet([['ID', 'Vehículo', 'Viaje', 'Criticidad', 'Fecha', 'Hora', 'Estado', 'Responsable', 'Turno'], ...rows]);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Alertas Jammer');
        XLSX.writeFile(workbook, filename);
    }

    document.addEventListener('DOMContentLoaded', () => {
        initCharts();
        document.getElementById('excelFile').addEventListener('change', event => { const file = event.target.files[0]; if (file) file.arrayBuffer().then(buffer => loadWorkbook(buffer, file.name)); });
        document.getElementById('selectFecha').addEventListener('change', () => { updateShifts(); render(); });
        document.getElementById('selectTurno').addEventListener('change', render);
        document.getElementById('selectFiltroEstado').addEventListener('change', render);
        document.getElementById('btnDescargarPendientes').addEventListener('click', () => exportExcel(datosFiltrados, `Alertas_Jammer_${new Date().toISOString().slice(0, 10)}.xlsx`));
        document.getElementById('btnExportPageExcel').addEventListener('click', () => exportExcel(datosFiltrados.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE), `Alertas_Jammer_Pagina_${currentPage}.xlsx`));
        document.getElementById('btnExportCharts').addEventListener('click', () => Object.entries(charts).forEach(([name, chart]) => { const link = document.createElement('a'); link.href = chart.toBase64Image(); link.download = `jammer_${name}.png`; link.click(); }));
        document.getElementById('sourceStatus').textContent = 'Seleccione un archivo Excel para comenzar.';
    });
}());
