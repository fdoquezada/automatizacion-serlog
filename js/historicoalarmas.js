let rawData = [];

document.getElementById('archivoExcel').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        rawData = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        console.log("Cargado");
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function procesar() {
    const filtro = document.getElementById('fechaFiltro').value;
    if (!filtro || rawData.length === 0) return alert("Carga un archivo y elige fecha");

    const cabecera = document.getElementById('cabeceraTabla');
    const cuerpo = document.getElementById('cuerpoTabla');
    
    const filtrados = rawData.filter(fila => {
        let valorFecha = fila["Fecha Alarma"] || fila["Fecha"] || "";
        if (!valorFecha) return false;

        let partes = valorFecha.toString().split(' ')[0].split('/');
        if(partes.length === 3) {
            let iso = `${partes[2]}-${partes[1].padStart(2, '0')}-${partes[0].padStart(2, '0')}`;
            return iso === filtro;
        }
        return false;
    });

    if (filtrados.length > 0) {
        const columnas = Object.keys(filtrados[0]);
        
        cabecera.innerHTML = `<tr>${columnas.map(c => `<th>${c}</th>`).join('')}</tr>`;
        cuerpo.innerHTML = filtrados.map(fila => 
            `<tr>${columnas.map(c => `<td>${fila[c]}</td>`).join('')}</tr>`
        ).join('');

        document.getElementById('contador').innerText = filtrados.length;
        document.getElementById('zonaResultados').style.display = 'block';
    } else {
        alert("No hay datos para esa fecha");
    }
}
