(function() {
    const fileInput = document.getElementById('file');
    const filePedidos = document.getElementById('filePedidos');
    const dateFilterInput = document.getElementById('dateFilter');
    const loadBtn = document.getElementById('loadBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const summaryTables = document.getElementById('summaryTables');
    const preview = document.getElementById('preview');
    const info = document.getElementById('info');

    let computed = null;
    let originalHeaderInfo = { org: 'agunsa', fecha: '' };

    // Mapa de servicios activos cargado desde el archivo Pedidos
    // Key: Referencia Externa (col B del pedidos), Value: Estado Pedido (col H)
    let serviciosActivosMap = {};

    // ── CARGA DEL ARCHIVO DE PEDIDOS (segundo archivo) ──────────────────────
    filePedidos.addEventListener('change', () => {
        const f = filePedidos.files[0];
        if (!f) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            // header:1 → array de arrays, primera fila son los headers
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            serviciosActivosMap = {};
            // Fila 0 = encabezados, datos desde fila 1
            // Col E (índice 4) = Verificacion (Toneladas)
            // Col H (índice 7) = Estado Pedido
            // Key para cruce: usamos col B (índice 1) Referencia Externa
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row || row.length === 0) continue;
                const refExterna = (row[1] || '').toString().trim();
                const toneladas  = (row[4] || '').toString().trim();   // Col E
                const estado     = (row[7] || '').toString().trim();   // Col H
                if (refExterna) {
                    // Si ya existe la key nos quedamos con el estado más relevante
                    // Prioridad: EN EJECUCION > SERVICIO FINALIZADO > ANULADO
                    const prev = serviciosActivosMap[refExterna];
                    const priority = { 'EN EJECUCION': 3, 'SERVICIO FINALIZADO': 2, 'ANULADO': 1 };
                    const pNew = priority[estado] || 0;
                    const pPrev = prev ? (priority[prev.estado] || 0) : -1;
                    if (pNew > pPrev) {
                        serviciosActivosMap[refExterna] = { toneladas, estado };
                    }
                }
            }

            const count = Object.keys(serviciosActivosMap).length;
            showFilePedidosStatus(`✓ ${count} registros cargados del archivo de pedidos.`);
        };
        reader.readAsArrayBuffer(f);
    });

    function showFilePedidosStatus(text) {
        const el = document.getElementById('pedidosStatus');
        if (el) { el.textContent = text; el.classList.remove('d-none'); }
    }

    function showInfo(text, isError) {
        info.textContent = text;
        info.className = `alert mt-3 ${isError ? 'alert-danger' : 'alert-info'}`;
        info.classList.remove('d-none');
    }

    // ── PROCESO PRINCIPAL ────────────────────────────────────────────────────
    loadBtn.addEventListener('click', () => {
        const f = fileInput.files[0];
        if (!f) return showInfo('Selecciona el archivo de eventos Gowit.', true);
        if (!dateFilterInput.value) return showInfo('Selecciona una fecha en el calendario.', true);

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            originalHeaderInfo.org   = rows[3] ? rows[3][5] || 'agunsa' : 'agunsa';
            originalHeaderInfo.fecha = rows[3] ? rows[3][6] || ''       : '';

            let headerIdx = 0;
            for (let i = 0; i < 15; i++) {
                if ((rows[i] || []).some(c => /riesgo|conductor/i.test(c))) { headerIdx = i; break; }
            }

            computed = processRows(rows[headerIdx], rows.slice(headerIdx + 1));
            renderCharts(computed);
            renderSummary(computed);
            renderPreview(computed.rowsDetail);
            downloadBtn.disabled = false;
            showInfo(`${computed.rowsDetail.length} eventos procesados. Fechas y riesgos normalizados.`, false);
        };
        reader.readAsArrayBuffer(f);
    });

    // ── PROCESAMIENTO DE FILAS ───────────────────────────────────────────────
    function processRows(header, dataRows) {
        const selectedDate = dateFilterInput.value;
        let refDate = new Date();
        if (originalHeaderInfo.fecha) {
            const p = originalHeaderInfo.fecha.split(' ')[0].split('/');
            refDate = new Date(p[2], p[1] - 1, p[0]);
        }

        const rowsDetail = [];
        const risk   = { ALTO: 0, MEDIO: 0, BAJO: 0 };
        const groups  = {};
        const types   = {};
        const hours   = Array.from({ length: 24 }, (_, i) => ({ key: `${i.toString().padStart(2, '0')}:00`, value: 0 }));

        dataRows.forEach(row => {
            let rowDateTxt  = (row[5] || '').toString();
            let finalDateStr = '';
            let timePart = rowDateTxt.includes(',') ? rowDateTxt.split(',')[1] : (rowDateTxt.split(' ')[1] || '');

            if (rowDateTxt.toLowerCase().includes('hoy')) {
                finalDateStr = refDate.toISOString().split('T')[0];
            } else if (rowDateTxt.toLowerCase().includes('ayer')) {
                let yesterday = new Date(refDate);
                yesterday.setDate(yesterday.getDate() - 1);
                finalDateStr = yesterday.toISOString().split('T')[0];
            } else {
                const parts = rowDateTxt.split(' ')[0].split('/');
                if (parts.length === 3) finalDateStr = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            }

            if (finalDateStr === selectedDate) {
                let rawRisk = (row[0] || '').toString().toUpperCase();
                let rsk = 'BAJO';
                if (rawRisk.includes('HIGH') || rawRisk.includes('ALTO'))       rsk = 'ALTO';
                else if (rawRisk.includes('MEDIUM') || rawRisk.includes('MEDIO')) rsk = 'MEDIO';

                risk[rsk]++;

                const g = row[3] || 'Sin Grupo';
                const t = row[4] || 'Otros';
                groups[g] = (groups[g] || 0) + 1;
                types[t]  = (types[t]  || 0) + 1;

                const m = rowDateTxt.match(/(\d{1,2}):/);
                if (m) hours[parseInt(m[1])].value++;

                const displayDate = finalDateStr.split('-').reverse().join('/') + (timePart ? ' ' + timePart.trim() : '');

                // ── CRUCE CON ARCHIVO DE PEDIDOS ─────────────────────────────
                // Usamos el campo Conductor (row[1]) o Grupo (row[3]) como clave de cruce.
                // Ajusta aquí la columna que corresponda a "Referencia Externa" en tu archivo Gowit.
                const conductor = (row[1] || '').toString().trim();
                const pedidoInfo = serviciosActivosMap[conductor] || null;

                let toneladas     = 'N/A';
                let servicioActivo = 'NO';

                if (pedidoInfo) {
                    toneladas = pedidoInfo.toneladas || 'N/A';
                    servicioActivo = pedidoInfo.estado === 'EN EJECUCION' ? 'SI — EN EJECUCION'
                                   : pedidoInfo.estado === 'SERVICIO FINALIZADO' ? pedidoInfo.estado
                                   : pedidoInfo.estado || 'NO';
                }

                rowsDetail.push({
                    Riesgo: rsk,
                    Conductor: conductor || 'N/A',
                    Vehiculo: row[2] || 'N/A',
                    Grupo: g,
                    Tipo: t,
                    Fecha: displayDate,
                    Toneladas: toneladas,
                    'Servicio Activo': servicioActivo
                });
            }
        });

        return {
            rowsDetail, risk, hours,
            groupsSorted: Object.entries(groups).sort((a, b) => b[1] - a[1]).map(e => ({ key: e[0], value: e[1] })),
            typesSorted:  Object.entries(types).sort((a, b) => b[1] - a[1]).map(e => ({ key: e[0], value: e[1] }))
        };
    }

    // ── RENDER TABLAS RESUMEN ────────────────────────────────────────────────
    function renderSummary(data) {
        summaryTables.innerHTML = '';
        const total = data.rowsDetail.length;
        const tables = [
            { t: 'Riesgos',          d: Object.entries(data.risk).map(e => ({ key: e[0], value: e[1] })) },
            { t: 'Top Grupos',       d: data.groupsSorted.slice(0, 5) },
            { t: 'Tipos de Eventos', d: data.typesSorted.slice(0, 5) }
        ];

        tables.forEach(sec => {
            const col = document.createElement('div');
            col.className = 'col-md-4';
            col.innerHTML = `<table class="table table-sm table-bordered table-red">
                <thead><tr><th>${sec.t}</th><th>Cant.</th></tr></thead>
                <tbody>${sec.d.map(i => `<tr><td>${i.key}</td><td>${i.value}</td></tr>`).join('')}
                <tr class="total-row"><td>TOTAL</td><td>${total}</td></tr></tbody></table>`;
            summaryTables.appendChild(col);
        });
    }

    // ── PREVIEW ──────────────────────────────────────────────────────────────
    function renderPreview(data) {
        preview.innerHTML = '';
        if (data.length === 0) return;
        const table = document.createElement('table');
        table.className = 'table table-sm table-striped';
        table.innerHTML = `<thead><tr>${Object.keys(data[0]).map(k => `<th>${k}</th>`).join('')}</tr></thead>
                           <tbody>${data.slice(0, 10).map(r => `<tr>${Object.values(r).map(v => {
                               let cls = '';
                               if (v === 'NO') cls = 'text-danger fw-bold';
                               else if (typeof v === 'string' && v.includes('EJECUCION')) cls = 'text-success fw-bold';
                               return `<td class="${cls}">${v}</td>`;
                           }).join('')}</tr>`).join('')}</tbody>`;
        preview.appendChild(table);
    }

    // ── DESCARGA EXCEL ───────────────────────────────────────────────────────
    downloadBtn.addEventListener('click', async () => {
        const workbook = new ExcelJS.Workbook();
        const resSheet = workbook.addWorksheet('RESUMEN EJECUTIVO');
        let curr = 1;

        const addRes = (title, data, chartId) => {
            resSheet.getCell(`A${curr}`).value = title;
            resSheet.getCell(`A${curr}`).font = { bold: true, size: 14, color: { argb: 'FFFF0000' } };
            curr++;
            resSheet.getRow(curr).values = ['Categoría', 'Cant.'];
            resSheet.getRow(curr).eachCell(c => {
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
                c.font = { color: { argb: 'FFFFFFFF' }, bold: true };
            });
            curr++;
            data.forEach(d => { resSheet.getRow(curr).values = [d.key, d.value]; curr++; });
            resSheet.getRow(curr).values = ['TOTAL GENERAL', computed.rowsDetail.length];
            resSheet.getRow(curr).font = { bold: true };
            resSheet.getRow(curr).eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } });

            const canvas = document.getElementById(chartId);
            if (canvas) {
                const img = workbook.addImage({ base64: canvas.toDataURL(), extension: 'png' });
                resSheet.addImage(img, { tl: { col: 4, row: curr - data.length - 2 }, ext: { width: 350, height: 180 } });
            }
            curr += 3;
        };

        addRes('1. NIVEL DE RIESGO',   Object.entries(computed.risk).map(e => ({ key: e[0], value: e[1] })), 'chartRisk');
        addRes('2. TOP 10 GRUPOS',     computed.groupsSorted.slice(0, 10), 'chartGroup');
        addRes('3. TIPOS DE EVENTOS',  computed.typesSorted.slice(0, 10),  'chartType');
        addRes('4. EVENTOS POR HORA',  computed.hours,                     'chartHours');

        // ── HOJA DETALLE ─────────────────────────────────────────────────────
        const detSheet = workbook.addWorksheet('DETALLE DE DATOS');
        detSheet.getCell('A1').value = 'gowit';
        detSheet.getCell('A1').font  = { size: 36, bold: true };
        detSheet.getCell('E1').value = 'Cantidad de eventos';
        detSheet.getCell('E2').value = computed.rowsDetail.length;
        detSheet.getCell('H1').value = 'Organización';
        detSheet.getCell('H2').value = originalHeaderInfo.org;
        detSheet.getCell('L1').value = 'Fecha de creación';
        detSheet.getCell('L2').value = originalHeaderInfo.fecha;

        ['E1', 'H1', 'L1'].forEach(c => {
            detSheet.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99FF33' } };
            detSheet.getCell(c).font = { bold: true };
        });

        const headerRow = detSheet.getRow(5);
        headerRow.values = Object.keys(computed.rowsDetail[0]);
        headerRow.eachCell(c => {
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99FF33' } };
            c.font = { bold: true };
        });

        computed.rowsDetail.forEach((r, i) => {
            const exRow = detSheet.getRow(i + 6);
            exRow.values = Object.values(r);
            // Color en columna "Servicio Activo" (última columna = índice 8, col I en 1-based)
            const colIdx = Object.keys(r).length; // último índice
            const cell   = exRow.getCell(colIdx);
            const val    = r['Servicio Activo'] || '';
            if (val === 'NO') {
                cell.font = { color: { argb: 'FFCC0000' }, bold: true };
            } else if (val.includes('EJECUCION')) {
                cell.font = { color: { argb: 'FF1a7f37' }, bold: true };
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe6f4ea' } };
            }
        });

        // Ajustar ancho de columnas en hoja Detalle
        detSheet.columns.forEach(col => { col.width = 22; });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), 'Reporte_Gowit_Completo.xlsx');
    });

    // ── GRÁFICAS ─────────────────────────────────────────────────────────────
    let charts = {};
    function renderCharts(data) {
        const draw = (id, type, labels, values, color) => {
            if (charts[id]) charts[id].destroy();
            charts[id] = new Chart(document.getElementById(id), {
                type,
                data: { labels, datasets: [{ label: 'Cantidad', data: values, backgroundColor: color }] },
                options: { animation: false, responsive: true }
            });
        };
        draw('chartRisk',   'doughnut', ['ALTO', 'MEDIO', 'BAJO'],             [data.risk.ALTO, data.risk.MEDIO, data.risk.BAJO], ['#dc3545', '#ffc107', '#198754']);
        draw('chartGroup',  'bar',      data.groupsSorted.slice(0, 5).map(x => x.key), data.groupsSorted.slice(0, 5).map(x => x.value), '#dc3545');
        draw('chartType',   'bar',      data.typesSorted.slice(0, 5).map(x => x.key),  data.typesSorted.slice(0, 5).map(x => x.value),  '#6c757d');
        draw('chartHours',  'line',     data.hours.map(x => x.key),             data.hours.map(x => x.value), '#dc3545');
    }
})();
