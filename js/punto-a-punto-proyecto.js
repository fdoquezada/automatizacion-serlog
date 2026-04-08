// ─────────────────────────────────────────────
//  CONFIGURACIÓN DE FILTROS
// ─────────────────────────────────────────────
const FILTROS = {
    colVerificacion: 4,   // Columna E (índice 0-based)
    valVerificacion: "TONELADA",
    colEstadoPedido: 7,   // Columna H
    valEstadoPedido: "EN EJECUCION",
    colEstadoViaje: 26    // Columna AA (índice 26)
};

// Columnas a mostrar en la tabla
const COLS_TABLA = [
    { idx: 0,  lbl: "Id Pedido" },
    { idx: 1,  lbl: "Ref. Externa" },
    { idx: 3,  lbl: "Categoría" },
    { idx: 4,  lbl: "Verificación" },
    { idx: 7,  lbl: "Estado Pedido" },
    { idx: 9,  lbl: "Cliente" },
    { idx: 14, lbl: "F. Plan Carga" },
    { idx: 18, lbl: "F. Plan Descarga" },
    { idx: 25, lbl: "Id Viaje" },
    { idx: 26, lbl: "Estado Viaje" },
    { idx: 27, lbl: "Nro DT" },
    { idx: 38, lbl: "Proveedor TTE" },
    { idx: 41, lbl: "Patente" },
    { idx: 43, lbl: "Chofer" },
    { idx: 47, lbl: "Estado Parada" },
    { idx: 48, lbl: "Tipo Parada" },
];

// ─────────────────────────────────────────────
//  ESTADO GLOBAL
// ─────────────────────────────────────────────
let rawData      = [];
let headers      = [];
let filteredData = [];
let displayData  = [];

// ─────────────────────────────────────────────
//  ELEMENTOS DOM
// ─────────────────────────────────────────────
const fileInput   = document.getElementById('fileInput');
const loadBtn     = document.getElementById('loadBtn');
const filterBtn   = document.getElementById('filterBtn');
const exportBtn   = document.getElementById('exportBtn');
const searchInput = document.getElementById('searchInput');
const alertBox    = document.getElementById('alertBox');
const statsBar    = document.getElementById('statsBar');
const tableArea   = document.getElementById('tableArea');
const backToTopBtn = document.getElementById('btn-back-to-top');

// ─────────────────────────────────────────────
//  HELPERS UI
// ─────────────────────────────────────────────
function showAlert(msg, type = 'info') {
    alertBox.className = `alert alert-${type} mt-3 py-2`;
    alertBox.innerHTML = msg;
    alertBox.classList.remove('d-none');
}
function hideAlert() { alertBox.classList.add('d-none'); }

if (backToTopBtn) {
    window.addEventListener('scroll', () => {
        backToTopBtn.style.display = window.scrollY > 300 ? 'flex' : 'none';
    });
    backToTopBtn.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
}

function normalize(val) {
    if (val === null || val === undefined) return "";
    return String(val).trim().toUpperCase();
}

// Badge por columna
function makeBadge(colIdx, val) {
    const v = normalize(val);
    if (colIdx === FILTROS.colVerificacion) {
        const cls = v === "TONELADA" ? "badge-tonelada" : "badge-larga";
        return `<span class="badge-pill ${cls}">${val || '—'}</span>`;
    }
    if (colIdx === FILTROS.colEstadoPedido) {
        let cls = "badge-finalizado";
        if (v.includes("EJECUCION")) cls = "badge-ejecucion";
        else if (v === "CREADO")     cls = "badge-creado";
        else if (v === "ANULADO")    cls = "badge-anulado";
        return `<span class="badge-pill ${cls}">${val || '—'}</span>`;
    }
    if (colIdx === FILTROS.colEstadoViaje) {
        let cls = "badge-viaje-finalizado";
        if (v === "ACTIVO")           cls = "badge-viaje-activo";
        else if (v === "INACTIVO")    cls = "badge-viaje-inactivo";
        else if (v.includes("RENDIR")) cls = "badge-viaje-rendir";
        else if (v === "ANULADO")      cls = "badge-viaje-anulado";
        return `<span class="badge-pill ${cls}">${val || '—'}</span>`;
    }
    return val !== null && val !== undefined && val !== "" ? String(val) : '<span style="color:#d0d5dd">—</span>';
}

// ─────────────────────────────────────────────
//  RENDERIZAR TABLA
// ─────────────────────────────────────────────
function renderTable(rows) {
    if (!rows || !rows.length) {
        tableArea.innerHTML = `
            <div class="empty-state">
                <i class="bi bi-inbox"></i>
                No hay filas que coincidan con los filtros aplicados
            </div>`;
        return;
    }

    const thead = COLS_TABLA.map(c => `<th>${c.lbl}</th>`).join('');
    const tbody = rows.map(row => {
        const cells = COLS_TABLA.map(c => `<td>${makeBadge(c.idx, row[c.idx])}</td>`).join('');
        return `<tr>${cells}</tr>`;
    }).join('');

    tableArea.innerHTML = `
        <div class="table-wrapper">
            <table class="dt">
                <thead><tr>${thead}</tr></thead>
                <tbody>${tbody}</tbody>
            </table>
        </div>
        <div class="px-3 py-2 text-muted" style="font-size:0.75rem; border-top:1px solid #f2f4f7;">
            Mostrando <b>${rows.length}</b> registro${rows.length !== 1 ? 's' : ''}
        </div>`;
}

// ─────────────────────────────────────────────
//  CARGAR EXCEL
// ─────────────────────────────────────────────
loadBtn.addEventListener('click', () => {
    if (!fileInput.files.length) {
        showAlert('<i class="bi bi-exclamation-triangle me-1"></i> Selecciona un archivo primero.', 'warning');
        return;
    }
    const file = fileInput.files[0];
    const reader = new FileReader();

    loadBtn.disabled = true;
    loadBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span> Cargando...';

    reader.onload = e => {
        try {
            const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

            if (allRows.length < 2) throw new Error("El archivo está vacío o no tiene datos.");

            headers = allRows[0];
            rawData = allRows.slice(1).filter(r => r.some(c => c !== "" && c !== undefined && c !== null));

            const cntTonelada  = rawData.filter(r => normalize(r[FILTROS.colVerificacion]) === FILTROS.valVerificacion).length;
            const cntEjecucion = rawData.filter(r => normalize(r[FILTROS.colEstadoPedido]) === FILTROS.valEstadoPedido).length;

            document.getElementById('statTotal').textContent     = rawData.length;
            document.getElementById('statTonelada').textContent  = cntTonelada;
            document.getElementById('statEjecucion').textContent = cntEjecucion;
            document.getElementById('statFiltrado').textContent  = '—';
            statsBar.classList.remove('d-none');

            filterBtn.disabled = false;
            showAlert(`<i class="bi bi-check-circle me-1"></i> Archivo cargado: <b>${rawData.length}</b> filas encontradas. Presiona <b>Filtrar</b> para aplicar los filtros.`, 'success');

            renderTable(rawData.slice(0, 50));

        } catch(err) {
            showAlert(`<i class="bi bi-x-circle me-1"></i> Error al leer el archivo: ${err.message}`, 'danger');
        } finally {
            loadBtn.disabled = false;
            loadBtn.innerHTML = '<i class="bi bi-upload me-1"></i> Cargar';
        }
    };

    reader.readAsArrayBuffer(file);
});

// ─────────────────────────────────────────────
//  FILTRAR (TONELADA + EN EJECUCION + ACTIVO/INACTIVO)
// ─────────────────────────────────────────────
filterBtn.addEventListener('click', () => {
    filteredData = rawData.filter(row => {
        const okVerif  = normalize(row[FILTROS.colVerificacion]) === FILTROS.valVerificacion;
        const okPedido = normalize(row[FILTROS.colEstadoPedido]) === FILTROS.valEstadoPedido;
        const estadoViaje = normalize(row[FILTROS.colEstadoViaje]);
        const okViaje = estadoViaje === 'ACTIVO' || estadoViaje === 'INACTIVO';
        return okVerif && okPedido && okViaje;
    });

    document.getElementById('statFiltrado').textContent = filteredData.length;
    displayData = [...filteredData];
    searchInput.value = '';
    renderTable(displayData);

    exportBtn.disabled = filteredData.length === 0;

    if (filteredData.length === 0) {
        showAlert(`<i class="bi bi-info-circle me-1"></i> Ninguna fila cumple los filtros: TONELADA + EN EJECUCION + (ACTIVO o INACTIVO).`, 'warning');
    } else {
        showAlert(`<i class="bi bi-funnel-fill me-1"></i> Filtro aplicado: <b>${filteredData.length}</b> filas coinciden con TONELADA + EN EJECUCION + (ACTIVO/INACTIVO).`, 'success');
    }
});

// ─────────────────────────────────────────────
//  BÚSQUEDA EN TIEMPO REAL
// ─────────────────────────────────────────────
searchInput.addEventListener('input', () => {
    const q = searchInput.value.trim().toLowerCase();
    const source = filteredData.length ? filteredData : rawData;
    if (!q) {
        displayData = [...source];
    } else {
        displayData = source.filter(row =>
            COLS_TABLA.some(c => String(row[c.idx] ?? '').toLowerCase().includes(q))
        );
    }
    renderTable(displayData);
});

// ─────────────────────────────────────────────
//  EXPORTAR EXCEL
// ─────────────────────────────────────────────
exportBtn.addEventListener('click', () => {
    if (!filteredData.length) return;

    const exportRows = [headers, ...filteredData];
    const ws = XLSX.utils.aoa_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Punto_a_Punto_Filtrado");

    const fecha = new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb, `Punto_a_Punto_Proyecto_${fecha}.xlsx`);
});
