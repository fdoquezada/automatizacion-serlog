/**
 * Módulo principal para el reporte de Pérdida de Señal - Tercer Aviso
 * Maneja carga de Excel, filtros, paginación, gráficos y exportaciones.
 */
(function() {
    'use strict';

    // ==================== UTILIDADES ====================

    const UTILS = {
        safeText: (value) => value == null ? '' : String(value).trim(),

        parseFechaEvento: (texto) => {
            if (!texto) return { fecha: '-', hora: '00:00', horaInt: 0 };
            if (texto instanceof Date && !isNaN(texto)) {
                const d = texto;
                const dia = String(d.getDate()).padStart(2, '0');
                const mes = String(d.getMonth() + 1).padStart(2, '0');
                const ano = String(d.getFullYear());
                const hora = String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
                return { fecha: `${ano}-${mes}-${dia}`, hora, horaInt: d.getHours() };
            }
            const s = String(texto).trim();
            const re = /([0-3]?\d)[\-/]([0-1]?\d)[\-/](\d{2,4})\s+(\d{1,2}):(\d{2})/;
            const m = s.match(re);
            if (m) {
                const dia = m[1].padStart(2, '0');
                const mes = m[2].padStart(2, '0');
                const ano = m[3].length === 2 ? '20' + m[3] : m[3];
                const hora = m[4].padStart(2, '0') + ':' + m[5];
                return { fecha: `${ano}-${mes}-${dia}`, hora, horaInt: parseInt(m[4], 10) };
            }
            return { fecha: '-', hora: s, horaInt: 0 };
        },

        esTratadaCampo: (valor) => {
            const s = UTILS.safeText(valor).toUpperCase();
            if (!s || s === '-') return false;
            if (/^\d+\s*d\s*\d+\s*h\s*\d+\s*m$/i.test(s)) return false;
            return true;
        },

        obtenerResponsable: (valor) => {
            const s = UTILS.safeText(valor);
            if (/^\d+\s*d\s*\d+\s*h\s*\d+\s*m$/i.test(s)) return '';
            if (s === '-' || s === '') return '';
            return s.toUpperCase();
        },

        obtenerTurno: (hora) => {
            let h = 0, m = 0;
            if (typeof hora === 'string') {
                const parts = hora.split(':');
                h = parseInt(parts[0], 10) || 0;
                m = parseInt(parts[1], 10) || 0;
            } else if (typeof hora === 'number') {
                h = hora;
            }
            const total = h * 60 + m;
            const inicioManana = 8 * 60;
            const finManana = 16 * 60;
            const inicioTarde = finManana + 1;
            const finTarde = 23 * 60 + 59;
            if (total >= inicioManana && total <= finManana) return 'Mañana';
            if (total >= inicioTarde && total <= finTarde) return 'Tarde';
            return 'Noche';
        },

        construirFechaVisual: (fechaYMD) => {
            if (!fechaYMD || fechaYMD === '-') return '-';
            const [ano, mes, dia] = fechaYMD.split('-');
            return `${dia}/${mes}/${ano}`;
        },

        obtenerConteoResponsables: (datos) => {
            return datos.reduce((acc, item) => {
                if (!item.responsable) return acc;
                acc[item.responsable] = (acc[item.responsable] || 0) + 1;
                return acc;
            }, {});
        },

        ordenarFechasDesc: (fechas) => {
            return fechas.sort((a, b) => {
                if (a === '-' || b === '-') return 0;
                return b.localeCompare(a);
            });
        }
    };

    // ==================== PROCESAMIENTO DE DATOS ====================

    function procesarFilas(rows) {
        const datos = [];
        const conteo = {};
        if (!Array.isArray(rows) || rows.length === 0) return { datos, conteo };

        const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
        const idx = key => headers.findIndex(h => h.includes(key));
        const iTipo = idx('tipo');
        const iId = idx('id');
        const iVeh = idx('veh');
        const iViaje = idx('viaj');
        const iTratada = idx('trat');
        const iEvento = idx('evento');
        const iCerrado = idx('cerrad');

        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row) continue;
            const tipo = UTILS.safeText(row[iTipo]);
            if (!/Perdida de Señal - Tercer Aviso/i.test(tipo)) continue;

            const id = UTILS.safeText(row[iId]);
            const vehiculo = UTILS.safeText(row[iVeh]);
            const viaje = UTILS.safeText(row[iViaje]);
            const tratada = row[iTratada];
            const evento = row[iEvento];
            const cerrado = iCerrado >= 0 ? row[iCerrado] : null;
            const eventoInfo = UTILS.parseFechaEvento(evento);
            const cerradoInfo = UTILS.parseFechaEvento(cerrado);
            const esTratada = UTILS.esTratadaCampo(tratada);
            const responsable = UTILS.obtenerResponsable(tratada);
            const turnoEvento = UTILS.obtenerTurno(eventoInfo.hora);
            const turnoGestion = esTratada ?
                (cerradoInfo.hora && cerradoInfo.hora !== '00:00' ? UTILS.obtenerTurno(cerradoInfo.hora) : turnoEvento) :
                '';

            if (responsable) {
                conteo[responsable] = (conteo[responsable] || 0) + 1;
            }

            datos.push({
                id,
                vehiculo,
                viaje,
                fechaEvento: eventoInfo.fecha,
                horaEvento: eventoInfo.hora,
                estado: esTratada ? 'Tratada' : 'Pendiente',
                esTratada,
                responsable: responsable || (esTratada ? 'Sin nombre' : 'Sin tratar'),
                turnoEvento,
                turnoGestion
            });
        }
        return { datos, conteo };
    }

    // ==================== RENDERIZADO ====================

    const RENDER = {
        resumen: (datosFiltrados) => {
            const total = datosFiltrados.length;
            const tratadas = datosFiltrados.filter(item => item.esTratada).length;
            const pendientes = datosFiltrados.filter(item => !item.esTratada).length;
            
            document.getElementById('txtTotalAlertas').textContent = total;
            document.getElementById('txtTotalTratadas').textContent = tratadas;
            document.getElementById('txtTotalPendientes').textContent = pendientes;
            
            // Actualizar gráficos de resumen
            actualizarGraficosResumen(total, tratadas, pendientes);
        },

        tabla: (datos, tbodyId = 'tablaAlertasBody') => {
            const tbody = document.getElementById(tbodyId);
            tbody.innerHTML = '';
            if (!Array.isArray(datos) || datos.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">No hay datos para mostrar</td></tr>';
                return;
            }
            datos.forEach(item => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${UTILS.safeText(item.id)}</td>
                    <td>${UTILS.safeText(item.vehiculo)}</td>
                    <td>${UTILS.safeText(item.viaje)}</td>
                    <td>${item.fechaEvento} ${item.horaEvento}</td>
                    <td><span class="badge ${item.esTratada ? 'bg-success' : 'bg-danger'}">${item.estado}</span></td>
                    <td>${UTILS.safeText(item.responsable)}</td>
                `;
                tbody.appendChild(tr);
            });
        },

        ranking: (datos, containerId = 'rankingTratadas') => {
            const container = document.getElementById(containerId);
            if (!container) return;
            const conteo = UTILS.obtenerConteoResponsables(datos);
            const list = Object.entries(conteo).sort((a, b) => b[1] - a[1]).slice(0, 5);
            if (list.length === 0) {
                container.innerHTML = '<div class="text-muted small">No hay registros de responsables.</div>';
                return;
            }
            container.innerHTML = list.map(([name, count], index) => `
                <div class="ranking-item">
                    <span><strong>#${index + 1} ${UTILS.safeText(name)}</strong></span>
                    <span class="badge bg-success">${count}</span>
                </div>
            `).join('');
        },

        detalleTurno: (datos, turno, fecha) => {
            const pendientes = datos.filter(item => 
                !item.esTratada && 
                item.turnoEvento === turno && 
                item.fechaEvento === fecha
            ).length;
            
            const tratadas = datos.filter(item => 
                item.esTratada && 
                item.turnoGestion === turno &&
                item.fechaEvento === fecha
            ).length;
            
            document.getElementById('txtPendientesTurno').textContent = pendientes;
            document.getElementById('txtGestionesTurno').textContent = tratadas;
            
            // Actualizar gráfico de turno
            actualizarGraficoTurno(pendientes, tratadas);
        }
    };

    // ==================== GRÁFICOS DE RESUMEN ====================

    let chartTotal = null;
    let chartEstadoResumen = null;
    let chartTurno = null;
    let chartRanking = null;

    function initGraficosResumen() {
        // Gráfico Total (Gauge circular)
        const ctxTotal = document.getElementById('chartTotalAlertas')?.getContext('2d');
        if (ctxTotal) {
            if (chartTotal) chartTotal.destroy();
            chartTotal = new Chart(ctxTotal, {
                type: 'doughnut',
                data: {
                    labels: ['Alertas'],
                    datasets: [{
                        data: [0, 1],
                        backgroundColor: ['#0d6efd', '#e9ecef'],
                        borderWidth: 0,
                        cutout: '75%'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false }
                    },
                    animation: {
                        animateRotate: true
                    }
                }
            });
        }

        // Gráfico Estado (Pendientes vs Tratadas)
        const ctxEstado = document.getElementById('chartEstadoResumen')?.getContext('2d');
        if (ctxEstado) {
            if (chartEstadoResumen) chartEstadoResumen.destroy();
            chartEstadoResumen = new Chart(ctxEstado, {
                type: 'doughnut',
                data: {
                    labels: ['Tratadas', 'Pendientes'],
                    datasets: [{
                        data: [0, 0],
                        backgroundColor: ['#198754', '#dc3545'],
                        borderWidth: 2,
                        borderColor: '#fff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { 
                            position: 'bottom',
                            labels: { font: { size: 9 }, boxWidth: 12 }
                        }
                    }
                }
            });
        }

        // Gráfico Turno (Pendientes vs Gestionadas)
        const ctxTurno = document.getElementById('chartTurnoResumen')?.getContext('2d');
        if (ctxTurno) {
            if (chartTurno) chartTurno.destroy();
            chartTurno = new Chart(ctxTurno, {
                type: 'doughnut',
                data: {
                    labels: ['Pendientes', 'Gestionadas'],
                    datasets: [{
                        data: [0, 0],
                        backgroundColor: ['#dc3545', '#198754'],
                        borderWidth: 2,
                        borderColor: '#fff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { 
                            position: 'bottom',
                            labels: { font: { size: 9 }, boxWidth: 12 }
                        }
                    }
                }
            });
        }
    }

    function actualizarGraficosResumen(total, tratadas, pendientes) {
        // Actualizar gráfico total
        if (chartTotal) {
            const maxTotal = Math.max(total, 1);
            chartTotal.data.datasets[0].data = [total, maxTotal];
            chartTotal.update();
        }

        // Actualizar gráfico estado
        if (chartEstadoResumen) {
            chartEstadoResumen.data.datasets[0].data = [tratadas, pendientes];
            chartEstadoResumen.update();
        }
    }

    function actualizarGraficoTurno(pendientes, gestionadas) {
        if (chartTurno) {
            chartTurno.data.datasets[0].data = [pendientes, gestionadas];
            chartTurno.update();
        }
    }

    // ==================== GRÁFICO RANKING ====================

    function initChartRanking() {
        const ctxR = document.getElementById('chartRanking')?.getContext('2d');
        if (ctxR) {
            if (chartRanking) chartRanking.destroy();
            chartRanking = new Chart(ctxR, {
                type: 'bar',
                data: {
                    labels: [],
                    datasets: [{
                        label: 'Tratadas',
                        data: [],
                        backgroundColor: '#0d6efd',
                        borderRadius: 4
                    }]
                },
                options: {
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: { display: false }
                    },
                    scales: {
                        x: { grid: { display: false } },
                        y: { grid: { display: false } }
                    }
                }
            });
        }
    }

    function actualizarChartRanking(datos) {
        if (!chartRanking) return;
        const conteo = UTILS.obtenerConteoResponsables(datos);
        const items = Object.entries(conteo).sort((a, b) => b[1] - a[1]).slice(0, 6);
        chartRanking.data.labels = items.map(i => i[0]);
        chartRanking.data.datasets[0].data = items.map(i => i[1]);
        chartRanking.update();
    }

    // ==================== PAGINACIÓN ====================

    let lastFilteredData = [];
    let currentPage = 1;
    const PAGE_SIZE = 25;

    function renderPage() {
        const start = (currentPage - 1) * PAGE_SIZE;
        const pageItems = lastFilteredData.slice(start, start + PAGE_SIZE);
        RENDER.tabla(pageItems);
        renderPagination(lastFilteredData.length);
    }

    function renderPagination(totalItems) {
        const container = document.getElementById('paginationControls');
        if (!container) return;
        const totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));

        let html = '';
        html += `<div class="btn-group btn-group-sm" role="group">`;
        html += `<button class="btn btn-outline-secondary" ${currentPage === 1 ? 'disabled' : ''} data-page="${currentPage - 1}">«</button>`;

        const maxButtons = 7;
        let startPage = Math.max(1, currentPage - Math.floor(maxButtons / 2));
        let endPage = Math.min(totalPages, startPage + maxButtons - 1);
        if (endPage - startPage + 1 < maxButtons) startPage = Math.max(1, endPage - maxButtons + 1);

        for (let p = startPage; p <= endPage; p++) {
            const active = p === currentPage ? 'btn-primary' : 'btn-outline-secondary';
            html += `<button class="btn ${active}" data-page="${p}">${p}</button>`;
        }

        html += `<button class="btn btn-outline-secondary" ${currentPage === totalPages ? 'disabled' : ''} data-page="${currentPage + 1}">»</button>`;
        html += `</div>`;
        html += `<span class="ms-2 text-muted small">${totalItems} registros</span>`;

        container.innerHTML = html;
        container.querySelectorAll('button[data-page]').forEach(btn => {
            btn.addEventListener('click', function() {
                const p = parseInt(this.getAttribute('data-page'), 10);
                if (p < 1 || p > totalPages) return;
                currentPage = p;
                renderPage();
                const scrollContainer = document.querySelector('.table-responsive.table-scroll');
                if (scrollContainer) scrollContainer.scrollTop = 0;
            });
        });
    }

    // ==================== FILTROS ====================

    function filtrarDatos(datos) {
        const fecha = document.getElementById('selectFecha').value;
        const turno = document.getElementById('selectTurno').value;
        const estado = document.getElementById('selectFiltroEstado').value;

        return datos.filter(item => {
            if (fecha && item.fechaEvento !== fecha) return false;
            if (turno) {
                const turnoActivo = item.esTratada ? item.turnoGestion : item.turnoEvento;
                if (turnoActivo !== turno) return false;
            }
            if (estado === 'pendientes' && item.esTratada) return false;
            if (estado === 'tratadas' && !item.esTratada) return false;
            return true;
        });
    }

    function aplicarFiltro(datos) {
        const fechaSeleccionada = document.getElementById('selectFecha').value;
        const turnoSeleccionado = document.getElementById('selectTurno').value;

        const filtroDatos = filtrarDatos(datos);
        RENDER.resumen(filtroDatos);

        if (!fechaSeleccionada) {
            RENDER.tabla([]);
            RENDER.ranking([]);
            document.getElementById('txtPendientesTurno').textContent = '0';
            document.getElementById('txtGestionesTurno').textContent = '0';
            actualizarChartRanking([]);
            return;
        }

        lastFilteredData = filtroDatos;
        currentPage = 1;
        renderPage();
        RENDER.ranking(filtroDatos);
        actualizarChartRanking(filtroDatos);

        if (turnoSeleccionado) {
            RENDER.detalleTurno(datos, turnoSeleccionado, fechaSeleccionada);
        } else {
            document.getElementById('txtPendientesTurno').textContent = '0';
            document.getElementById('txtGestionesTurno').textContent = '0';
        }
    }

    // ==================== EXPORTACIONES ====================

    function downloadURI(uri, name) {
        const link = document.createElement('a');
        link.href = uri;
        link.download = name;
        document.body.appendChild(link);
        link.click();
        link.remove();
    }

    function descargarPendientesExcel(datos) {
        const filas = filtrarDatos(datos);
        if (!filas.length) {
            alert('No hay datos para exportar');
            return;
        }
        const encabezados = ['ID', 'Vehículo', 'Viaje', 'Fecha', 'Hora', 'Estado', 'Responsable', 'Turno'];
        const datosExport = filas.map(item => [
            item.id,
            item.vehiculo,
            item.viaje,
            UTILS.construirFechaVisual(item.fechaEvento),
            item.horaEvento,
            item.estado,
            item.responsable,
            item.esTratada ? item.turnoGestion : item.turnoEvento
        ]);
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([encabezados, ...datosExport]);
        XLSX.utils.book_append_sheet(wb, ws, 'Pendientes');
        const hoy = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `Pendientes_PerdidaSenal_${hoy}.xlsx`);
    }

    function exportCurrentPageExcel() {
        const start = (currentPage - 1) * PAGE_SIZE;
        const pageItems = lastFilteredData.slice(start, start + PAGE_SIZE);
        if (!pageItems.length) {
            alert('No hay filas en la página actual para exportar.');
            return;
        }
        const encabezados = ['ID', 'Vehículo', 'Viaje', 'Fecha', 'Hora', 'Estado', 'Responsable', 'Turno'];
        const datosExport = pageItems.map(item => [
            item.id,
            item.vehiculo,
            item.viaje,
            UTILS.construirFechaVisual(item.fechaEvento),
            item.horaEvento,
            item.estado,
            item.responsable,
            item.esTratada ? item.turnoGestion : item.turnoEvento
        ]);
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([encabezados, ...datosExport]);
        XLSX.utils.book_append_sheet(wb, ws, 'Pagina');
        const hoy = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `PerdidaSenal_Pagina_${hoy}_p${currentPage}.xlsx`);
    }

    function exportChartsPNG() {
        const charts = [chartTotal, chartEstadoResumen, chartTurno, chartRanking].filter(c => c);
        if (!charts.length) {
            alert('No hay gráficos disponibles para exportar.');
            return;
        }
        charts.forEach((chart, index) => {
            const names = ['total', 'estado', 'turno', 'ranking'];
            const img = chart.toBase64Image();
            downloadURI(img, `chart_${names[index] || index}_${new Date().toISOString().slice(0, 10)}.png`);
        });
    }

    // ==================== INICIALIZACIÓN ====================

    document.addEventListener('DOMContentLoaded', function() {
        const excelFile = document.getElementById('excelFile');
        const selectFecha = document.getElementById('selectFecha');
        const selectTurno = document.getElementById('selectTurno');
        const filtroEstado = document.getElementById('selectFiltroEstado');

        let datosGlobal = [];

        // === FILTROS RÁPIDOS EN CABECERA ===
        document.querySelectorAll('.btn-filtro-estado').forEach(btn => {
            btn.addEventListener('click', function() {
                // Remover active de todos
                document.querySelectorAll('.btn-filtro-estado').forEach(b => b.classList.remove('active'));
                this.classList.add('active');
                
                const estado = this.getAttribute('data-estado');
                filtroEstado.value = estado;
                aplicarFiltro(datosGlobal);
            });
        });

        function actualizarTurnosDisponibles() {
            const fecha = selectFecha.value;
            const baseTurnos = ['Mañana', 'Tarde', 'Noche'];
            const currentTurno = selectTurno.value;
            
            const opciones = baseTurnos.map(t => {
                const count = datosGlobal.filter(item => {
                    if (fecha && item.fechaEvento !== fecha) return false;
                    return (item.turnoEvento === t) || (item.turnoGestion === t);
                }).length;
                const selected = (t === currentTurno) ? 'selected' : '';
                return `<option value="${t}" ${selected}>${t} ${count ? '(' + count + ')' : ''}</option>`;
            }).join('');
            
            selectTurno.innerHTML = '<option value="">Seleccione un turno</option>' + opciones;
            selectTurno.disabled = !selectFecha.value;
        }

        function actualizarTodo() {
            if (datosGlobal.length === 0) return;
            aplicarFiltro(datosGlobal);
        }

        // Event listeners
        filtroEstado.addEventListener('change', actualizarTodo);
        
        selectFecha.addEventListener('change', function() {
            actualizarTurnosDisponibles();
            const turnoActual = selectTurno.value;
            if (turnoActual) {
                const opciones = selectTurno.querySelectorAll('option');
                let encontrado = false;
                for (let opt of opciones) {
                    if (opt.value === turnoActual) {
                        encontrado = true;
                        break;
                    }
                }
                if (!encontrado && opciones.length > 1) {
                    selectTurno.value = opciones[1].value;
                }
            }
            actualizarTodo();
        });
        
        selectTurno.addEventListener('change', actualizarTodo);

        // Botones de exportación
        document.getElementById('btnDescargarPendientes').addEventListener('click', () => {
            descargarPendientesExcel(datosGlobal);
        });
        document.getElementById('btnExportPageExcel').addEventListener('click', exportCurrentPageExcel);
        document.getElementById('btnExportCharts').addEventListener('click', exportChartsPNG);

        // Carga de archivo Excel
        excelFile.addEventListener('change', function(event) {
            const file = event.target.files && event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                const sheetName = workbook.SheetNames[0];
                const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

                const resultado = procesarFilas(rows);
                datosGlobal = resultado.datos;

                const fechas = Array.from(new Set(datosGlobal.map(item => item.fechaEvento)
                    .filter(f => f && f !== '-'))).sort((a, b) => b.localeCompare(a));
                
                selectFecha.innerHTML = '<option value="">Seleccione un día</option>' +
                    fechas.map(f => `<option value="${f}">${UTILS.construirFechaVisual(f)}</option>`).join('');

                selectFecha.disabled = false;
                filtroEstado.disabled = false;
                selectTurno.disabled = true;

                if (fechas.length > 0) {
                    selectFecha.value = fechas[0];
                    actualizarTurnosDisponibles();
                    
                    const opciones = selectTurno.querySelectorAll('option');
                    let primerTurno = '';
                    for (let opt of opciones) {
                        if (opt.value && !opt.value.includes('(')) {
                            primerTurno = opt.value;
                            break;
                        }
                    }
                    if (primerTurno) {
                        selectTurno.value = primerTurno;
                    }
                    selectTurno.disabled = false;
                }

                // Inicializar gráficos
                initGraficosResumen();
                initChartRanking();

                actualizarTodo();
            };
            reader.readAsArrayBuffer(file);
        });

        // Inicializar gráficos al cargar la página
        initGraficosResumen();
        initChartRanking();

        // Botón volver arriba
        const backToTop = document.getElementById('btn-back-to-top');
        if (backToTop) {
            backToTop.addEventListener('click', () => {
                window.scrollTo({ top: 0, behavior: 'smooth' });
            });
            window.addEventListener('scroll', () => {
                backToTop.style.display = window.scrollY > 300 ? 'block' : 'none';
            });
            backToTop.style.display = 'none';
        }
    });
})();