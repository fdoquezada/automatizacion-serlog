// Mapeo de nombres de columnas para estandarización
const COLUMN_MAPPING = {
    'Categoria': 'Categoría',
    'Estado Viaje': 'Estado Viaje',
    'Estado': 'Estado',
    'Estado Pedido': 'Estado Pedido',
    'EstadoPedido': 'Estado Pedido',
    'Estado Pedido ': 'Estado Pedido',
    'Estado_Pedido': 'Estado Pedido',
    'Transporte Antes de Hora Plan': 'Transporte Antes Hora Plan',
    'TransporteAntesDeHoraPlan': 'Transporte Antes Hora Plan',
    'Transporte Antes de Hora (Plan)': 'Transporte Antes Hora Plan',
    'Transporte': 'Transporte Antes Hora Plan',
    'Cliente': 'Cliente',
    'Comuna': 'Comuna',
    'Patente': 'Patente'
};

function normalizarNombreColumna(nombre) {
    return COLUMN_MAPPING[nombre] || nombre;
}

function normalizarTexto(texto) {
    return String(texto || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

const app = {
    datosGlobales: [],
    lastFilteredData: [],
    actualFiltro: 'TOTAL',
    currentColumns: [],
    sortConfig: { col: null, asc: true },
    
    limpiarTodo() {
        this.datosGlobales = [];
        this.lastFilteredData = [];
        this.actualFiltro = 'TOTAL';
        this.currentColumns = [];
        this.sortConfig = { col: null, asc: true };
        
        document.getElementById('excelFile').value = '';
        document.getElementById('searchText').value = '';
        document.getElementById('columnaSelect').innerHTML = '<option value="ANY">Todas las columnas</option>';
        
        document.getElementById('statsBar').style.display = 'none';
        document.getElementById('tableSection').style.display = 'none';
        
        document.getElementById('thead').innerHTML = '';
        document.getElementById('tbody').innerHTML = '';
        
        document.getElementById('countTotal').innerText = '0';
        document.getElementById('countActivo').innerText = '0';
        document.getElementById('countInactivo').innerText = '0';
        
        alert('Datos limpiados correctamente. Puede cargar un nuevo archivo Excel.');
    },
    
    procesarExcel() {
        const file = document.getElementById('excelFile').files[0];
        if (!file) return alert("Seleccione un archivo");

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            
            this.datosGlobales = json.map(fila => {
                const filaNormalizada = {};
                Object.keys(fila).forEach(key => {
                    filaNormalizada[normalizarNombreColumna(key)] = fila[key];
                });
                return filaNormalizada;
            });
            
            this.calcularResumen(this.datosGlobales);
        };
        reader.readAsArrayBuffer(file);
    },
    
    calcularResumen(data) {
        let activos = 0, inactivos = 0;

        data.forEach(fila => {
            const cat = normalizarTexto(fila['Categoría'] || '');
            const estado = normalizarTexto(fila['Estado Viaje'] || fila['Estado'] || '');

            if (cat.includes('urgencia')) {
                if (estado === 'activo') activos++;
                if (estado === 'inactivo') inactivos++;
            }
        });

        const totalReal = activos + inactivos;

        document.getElementById('countTotal').innerText = totalReal;
        document.getElementById('countActivo').innerText = activos;
        document.getElementById('countInactivo').innerText = inactivos;
        document.getElementById('statsBar').style.display = 'flex';
        
        this.aplicarFiltro('TOTAL');
    },
    
    aplicarFiltro(tipo) {
        this.actualFiltro = tipo;
        const txtVista = document.getElementById('txtVista');
        const badge = document.getElementById('badgeConteo');

        const filtrados = this.datosGlobales.filter(fila => {
            const cat = normalizarTexto(fila['Categoría'] || '');
            const estado = normalizarTexto(fila['Estado Viaje'] || fila['Estado'] || '');
            
            if (!cat.includes('urgencia')) return false;

            if (tipo === 'ACTIVO') return estado === 'activo';
            if (tipo === 'INACTIVO') return estado === 'inactivo';
            return true;
        });

        const texto = String(document.getElementById('searchText').value || '').trim().toLowerCase();
        const columna = document.getElementById('columnaSelect').value;

        const filtradosConBusqueda = filtrados.filter(fila => {
            if (!texto) return true;
            if (columna && columna !== 'ANY') {
                const valor = normalizarTexto(fila[columna] || '');
                return valor.includes(texto);
            }
            return Object.values(fila).some(v => normalizarTexto(v || '').includes(texto));
        });

        this.lastFilteredData = filtradosConBusqueda;
        txtVista.innerHTML = `<i class="bi bi-filter"></i> Urgencias ${tipo === 'TOTAL' ? 'Todas' : tipo}`;
        badge.innerText = `${filtradosConBusqueda.length} registros`;
        this.renderizarTabla(filtradosConBusqueda);
        document.getElementById('tableSection').style.display = 'block';
    },
    
    renderizarTabla(datos) {
        const thead = document.getElementById('thead');
        const tbody = document.getElementById('tbody');
        thead.innerHTML = '';
        tbody.innerHTML = '';

        if (datos.length === 0) {
            tbody.innerHTML = '<tr><td colspan="100%" class="text-center">No hay datos para mostrar</td></tr>';
            return;
        }

        const columnas = Object.keys(datos[0]);
        this.currentColumns = columnas;
        this.poblarColumnas(columnas);

        const headerRow = '<tr>' + columnas.map(c => 
            `<th style="cursor:pointer; white-space:nowrap;" onclick="app.toggleSort('${c.replace(/'/g, "\\'")}')">
                ${c} 
                <span class="sort-indicator" id="sort_${c.replace(/\s/g, '_')}">
                    ${this.sortConfig.col === c ? (this.sortConfig.asc ? '▲' : '▼') : ''}
                </span>
            </th>`
        ).join('') + '</tr>';
        
        thead.innerHTML = headerRow;

        let mostrados = [...datos];
        if (this.sortConfig.col) {
            mostrados = mostrados.sort((a, b) => {
                const va = normalizarTexto(a[this.sortConfig.col] || '');
                const vb = normalizarTexto(b[this.sortConfig.col] || '');
                if (va < vb) return this.sortConfig.asc ? -1 : 1;
                if (va > vb) return this.sortConfig.asc ? 1 : -1;
                return 0;
            });
        }

        mostrados.forEach(fila => {
            const tr = document.createElement('tr');
            columnas.forEach(c => {
                const td = document.createElement('td');
                td.textContent = fila[c] || '';
                td.style.whiteSpace = 'nowrap';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    },
    
    toggleSort(col) {
        if (this.sortConfig.col === col) {
            this.sortConfig.asc = !this.sortConfig.asc;
        } else {
            this.sortConfig.col = col;
            this.sortConfig.asc = true;
        }
        this.renderizarTabla(this.lastFilteredData);
    },
    
    poblarColumnas(columnas) {
        const sel = document.getElementById('columnaSelect');
        sel.innerHTML = '<option value="ANY">Todas las columnas</option>' + 
            columnas.map(c => `<option value="${c.replace(/"/g, '&quot;')}">${c}</option>`).join('');
    },
    
    exportarExcelNuevo() {
        if (!this.lastFilteredData || this.lastFilteredData.length === 0) {
            return alert('No hay datos para exportar');
        }

        const datosFiltradosParaExport = this.lastFilteredData.filter(fila => {
            const estadoPedido = normalizarTexto(fila['Estado Pedido'] || fila['Estado'] || '');
            const transporte = normalizarTexto(fila['Transporte Antes Hora Plan'] || '');

            const esEnEjecucion = estadoPedido.includes('en ejec') || 
                                 estadoPedido === 'enejecucion' || 
                                 estadoPedido.includes('en ejecucion');
            const esSi = transporte === 'si' || transporte === 's' || transporte.includes('si');

            return esEnEjecucion && esSi;
        });

        if (datosFiltradosParaExport.length === 0) {
            return alert('No hay registros que cumplan: Estado Pedido = "en ejecucion" y Transporte Antes de Hora Plan = "si"');
        }

        const sheet = XLSX.utils.json_to_sheet(datosFiltradosParaExport);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, sheet, 'Urgencias_Filtradas');
        
        const ahora = new Date();
        const fecha = `${ahora.getFullYear()}-${String(ahora.getMonth()+1).padStart(2,'0')}-${String(ahora.getDate()).padStart(2,'0')}_${String(ahora.getHours()).padStart(2,'0')}${String(ahora.getMinutes()).padStart(2,'0')}`;
        XLSX.writeFile(wb, `urgencias_export_${fecha}.xlsx`);
    }
};

window.app = app;
