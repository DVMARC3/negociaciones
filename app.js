let appData = []; // Aquí guardaremos los datos del Excel
const fileUpload = document.getElementById('fileUpload');
const exportBtn = document.getElementById('exportBtn');
const tableHeader = document.getElementById('tableHeader');
const tableBody = document.getElementById('tableBody');

// Definimos el orden exacto de las columnas de tu Excel
const columnasExcel = [
    "PROVINCIA", "SERVICIO CONCERTADO", "ESPECIALIDAD", "CÓDIGO ESPECIALIDAD", 
    "GRUPO HOSPITALARIO", "COD GH", "CÓDIGO SC", "IMPORTE", "SUBIDA FIJA 2027", "BOLSA",
    "% NEGOCIADO", "BOLSA NEGOCIACIÓN", "JUSTIFICACIÓN % PROPUESTO", "SUBIDA LINEAL", 
    "REVISADO POR RESPONSABLE", "OBSERVACIONES"
];

// Definimos cuáles son editables por el negociador
const columnasEditables = [
    "% NEGOCIADO", "BOLSA NEGOCIACIÓN", "JUSTIFICACIÓN % PROPUESTO", 
    "SUBIDA LINEAL", "REVISADO POR RESPONSABLE", "OBSERVACIONES"
];

// 1. Leer archivo Excel
fileUpload.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Asumimos que los datos están en la primera hoja
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convertimos a JSON
        appData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        renderTable();
        exportBtn.disabled = false;
    };
    reader.readAsArrayBuffer(file);
});

// 2. Renderizar la tabla en HTML
function renderTable() {
    // Generar Cabeceras
    tableHeader.innerHTML = '';
    columnasExcel.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        tableHeader.appendChild(th);
    });

    // Generar Filas
    tableBody.innerHTML = '';
    appData.forEach((fila, index) => {
        const tr = document.createElement('tr');
        
        columnasExcel.forEach(col => {
            const td = document.createElement('td');
            const valorActual = fila[col] !== undefined ? fila[col] : "";

            if (columnasEditables.includes(col)) {
                // Crear input para campos editables
                td.classList.add('editable');
                const input = document.createElement('input');
                input.type = 'text';
                input.value = valorActual;
                // Escuchar cambios para actualizar appData
                input.addEventListener('change', (e) => {
                    appData[index][col] = e.target.value;
                });
                td.appendChild(input);
            } else {
                // Texto estático para campos de solo lectura
                td.textContent = valorActual;
            }
            tr.appendChild(td);
        });
        
        tableBody.appendChild(tr);
    });
}

// 3. Guardar y exportar el Excel modificado
exportBtn.addEventListener('click', () => {
    if (appData.length === 0) return;

    // Crear nueva hoja de cálculo con los datos modificados
    const worksheet = XLSX.utils.json_to_sheet(appData, { header: columnasExcel });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Negociaciones");

    // Descargar el archivo
    XLSX.writeFile(workbook, "Negociaciones_Actualizadas.xlsx");
});