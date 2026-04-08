// JS: feedback visual al seleccionar archivos
function setupFileBox(inputId, boxId, statusId) {
    const input  = document.getElementById(inputId);
    const box    = document.getElementById(boxId);
    const status = document.getElementById(statusId);
    input.addEventListener('change', () => {
        if (input.files[0]) {
            box.classList.add('has-file');
            status.classList.remove('d-none');
            status.querySelector('span').textContent = input.files[0].name;
        } else {
            box.classList.remove('has-file');
            status.classList.add('d-none');
        }
    });
}
setupFileBox('file',        'boxGowit',   'gowitStatus');
setupFileBox('filePedidos', 'boxPedidos', 'pedidosStatus');
