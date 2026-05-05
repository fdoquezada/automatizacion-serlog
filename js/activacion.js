let resultados = { activar: [], cerrar: [] };
let filtrosTabla = {
    tablaActivar: { idViaje:'', cliente:'', fecha:'', origen:'', texto:'' },
    tablaCerrar: { idViaje:'', cliente:'', fecha:'', origen:'', texto:'' }
};
let sortState = {
    tablaActivar: { field: null, dir: 1 },
    tablaCerrar: { field: null, dir: 1 }
};

window.onload = () => {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    document.getElementById('fechaInforme').value = now.toISOString().slice(0,16);
    initTablaFiltros();
    initTablaOrden();
};

function leer(file, esCap) {
    if(!file) return alert("Seleccione un archivo");
    const reader = new FileReader();
    reader.onload = (e) => {
        const wb = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
        let data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:""});
        data = data.map(r => {
            let c = {};
            Object.keys(r).forEach(k => c[k.trim().replace(/[^\x00-\x7F]/g, "")] = r[k]);
            return c;
        });
        ejecutarLogica(data, esCap);
    };
    reader.readAsArrayBuffer(file);
}

function procesarNormal() { 
    leer(document.getElementById('fileNormal').files[0], false); 
}

function procesarCapstone() { 
    leer(document.getElementById('fileCapstone').files[0], true); 
}

function parseFecha(v) {
    if(!v) return null;
    if(typeof v === 'number') return new Date(Math.round((v-25569)*86400*1000));
    const p = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s*(\d{1,2}):(\d{1,2})/);
    return p ? new Date(p[3], p[2]-1, p[1], p[4], p[5]) : (isNaN(new Date(v).getTime())?null:new Date(v));
}

function ejecutarLogica(data, esCap) {
    const ref = new Date(document.getElementById('fechaInforme').value);
    const limite = new Date(ref.getTime() - (60*60*1000));
    let nuevosEncontrados = 0;
    
    data.forEach(f => {
        const idV = f['Id Viaje'];
        if(!idV) return;

        if(esCap) {
            if(!String(f['Verificacion']||f['Verificación']||"").toUpperCase().includes("TONELADA")) return;
            const creador = String(f['Usuario Creador']||"").toLowerCase();
            if(creador !== 'acalderon' && creador !== 'yrosales') return;
        }

        const est = String(f['Estado Viaje']||"").toLowerCase();
        const cat = String(f['Categoria']||"").toLowerCase();

        if((est==='activo'||est==='inactivo') && !cat.includes('urgencia')) {
            const ori = esCap ? "CAPSTONE" : "PTO A PTO";
            const oriClass = esCap ? "col-origen-cap" : "col-origen-pto";

            if(est==='inactivo') {
                const fi = parseFecha(f['Viaje Fecha Inicio Plan']);
                if(fi && fi <= limite && !resultados.activar.some(x=>x.idViaje===idV)) {
                    resultados.activar.push({idViaje:idV, cliente:f['Cliente'], fecha:fi, texto:'POR ACTIVAR', origen:ori, oClass:oriClass});
                    nuevosEncontrados++;
                }
            } else if(est==='activo') {
                const ff = parseFecha(f['Viaje Fecha Fin Plan']);
                if(ff && ff <= limite && !resultados.cerrar.some(x=>x.idViaje===idV)) {
                    resultados.cerrar.push({idViaje:idV, cliente:f['Cliente'], fecha:ff, texto:'GESTIONAR CIERRE', origen:ori, oClass:oriClass});
                    nuevosEncontrados++;
                }
            }
        }
    });

    if (nuevosEncontrados > 0) {
        alert(`✅ Se procesaron los datos. Se encontraron ${nuevosEncontrados} viajes pendientes.`);
    } else {
        alert("⚠️ No se encontraron viajes que cumplan con los filtros de tiempo, usuario o categoría en este archivo.");
    }

    document.getElementById('containerDescarga').style.display = "block";
    render();
}

function initTablaFiltros() {
    document.querySelectorAll('.table-filter-input').forEach(input => {
        input.addEventListener('input', () => {
            const tableId = input.dataset.table;
            const field = input.dataset.field;
            filtrosTabla[tableId][field] = input.value.trim().toLowerCase();
            render();
        });
    });
}

function initTablaOrden() {
    document.querySelectorAll('th.sortable').forEach(th => {
        th.style.cursor = 'pointer';
        th.addEventListener('click', () => {
            const tableId = th.dataset.table;
            const field = th.dataset.field;
            toggleOrden(tableId, field);
            render();
        });
    });
    actualizarIndicadoresOrden();
}

function toggleOrden(tableId, field) {
    if (sortState[tableId].field === field) {
        sortState[tableId].dir = sortState[tableId].dir * -1;
    } else {
        sortState[tableId].field = field;
        sortState[tableId].dir = 1;
    }
    actualizarIndicadoresOrden();
}

function actualizarIndicadoresOrden() {
    document.querySelectorAll('th.sortable').forEach(th => {
        const tableId = th.dataset.table;
        const field = th.dataset.field;
        const icon = th.querySelector('.sort-icon');
        if (!icon) return;
        if (sortState[tableId].field === field) {
            icon.textContent = sortState[tableId].dir === 1 ? '▲' : '▼';
        } else {
            icon.textContent = '';
        }
    });
}

function ordenarDatos(arr, tableId) {
    const { field, dir } = sortState[tableId];
    if (!field) return arr.slice();
    return arr.slice().sort((a, b) => {
        let va = a[field];
        let vb = b[field];
        if (va instanceof Date && vb instanceof Date) {
            return (va - vb) * dir;
        }
        va = va == null ? '' : String(va).toLowerCase();
        vb = vb == null ? '' : String(vb).toLowerCase();
        return va.localeCompare(vb, undefined, { numeric: true }) * dir;
    });
}

function formatearFecha(fecha) {
    const dia = fecha.getDate().toString().padStart(2, '0');
    const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
    const anio = fecha.getFullYear();
    const hora = fecha.getHours().toString().padStart(2, '0');
    const minuto = fecha.getMinutes().toString().padStart(2, '0');
    const segundo = fecha.getSeconds().toString().padStart(2, '0');
    return `${dia}-${mes}-${anio}, ${hora}:${minuto}:${segundo}`;
}

function aplicarFiltros(arr, tableId) {
    const filtro = filtrosTabla[tableId];
    return arr.filter(d => {
        const fechaTexto = formatearFecha(d.fecha).toLowerCase();
        return (!filtro.idViaje || String(d.idViaje).toLowerCase().includes(filtro.idViaje)) &&
               (!filtro.cliente || String(d.cliente).toLowerCase().includes(filtro.cliente)) &&
               (!filtro.fecha || fechaTexto.includes(filtro.fecha)) &&
               (!filtro.origen || String(d.origen).toLowerCase().includes(filtro.origen)) &&
               (!filtro.texto || String(d.texto).toLowerCase().includes(filtro.texto));
    });
}

function render() {
    const draw = (arr, sel, accClass) => {
        const tableId = sel.replace('#', '');
        const sorted = ordenarDatos(arr, tableId);
        const filtered = aplicarFiltros(sorted, tableId);
        document.querySelector(`${sel} tbody`).innerHTML = filtered.map(d => `
            <tr>
                <td class="text-center"><b>${d.idViaje}</b></td>
                <td>${d.cliente}</td>
                <td class="text-center">${formatearFecha(d.fecha)}</td>
                <td class="${d.oClass}">${d.origen}</td>
                <td class="${accClass}">${d.texto}</td>
            </tr>`).join('') || '<tr><td colspan="5" class="text-center py-3 text-muted">Sin datos</td></tr>';
    };
    draw(resultados.activar, '#tablaActivar', 'col-accion-activar');
    draw(resultados.cerrar, '#tablaCerrar', 'col-accion-cierre');
    document.getElementById('countActivar').innerText = aplicarFiltros(ordenarDatos(resultados.activar, 'tablaActivar'), 'tablaActivar').length;
    document.getElementById('countCerrar').innerText = aplicarFiltros(ordenarDatos(resultados.cerrar, 'tablaCerrar'), 'tablaCerrar').length;
}

function descargarExcel() {
    const wb = XLSX.utils.book_new();
    const sheet = (data) => {
        const h = [["ID Viaje", "Cliente", "Fecha", "Origen", "Estado"]];
        const b = data.map(d => [d.idViaje, d.cliente, formatearFecha(d.fecha), d.origen, d.texto]);
        const ws = XLSX.utils.aoa_to_sheet(h.concat(b));
        ws['!cols'] = [{wch:12}, {wch:35}, {wch:22}, {wch:15}, {wch:20}];
        return ws;
    };
    if(resultados.activar.length) XLSX.utils.book_append_sheet(wb, sheet(resultados.activar), "Por Activar");
    if(resultados.cerrar.length) XLSX.utils.book_append_sheet(wb, sheet(resultados.cerrar), "Gestionar Cierre");
    XLSX.writeFile(wb, `Reporte_Agunsa_${new Date().toLocaleDateString()}.xlsx`);
}
