document.addEventListener("DOMContentLoaded", () => {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";
  
    const plantas = {
      "GU01": "VILLA NUEVA",
      "GU02": "SALCAJA",
      "GU03": "HUEHUETENANGO",
      "GU04": "ZACAPA",
      "GU05": "TIKAL",
      "GU06": "COBAN",
      "GU07": "QUICHE",
      "GU08": "TOTONICAPAN",
      "GU09": "ESCUINTLA",
      "GU10": "JUTIAPA"
    };
  
    const columnas = [
      "FOLIO", "FECHA", "SERIE", "FACTURA",
      "VILLA NUEVA", "SALCAJA", "HUEHUETENANGO", "ZACAPA", "TIKAL",
      "COBAN", "QUICHE", "TOTONICAPAN", "ESCUINTLA", "JUTIAPA",
      "VALOR GTQ", "VALOR EN DLL", "GALONES TOTALES"
    ];
  
    let registros = [];
    let duplicados = [];
  
    // Crear encabezado dinámico
    const headerRow = document.getElementById("headerRow");
    columnas.forEach(col => {
      const th = document.createElement("th");
      th.textContent = col;
      headerRow.appendChild(th);
    });
  
    document.getElementById("pdfFiles").addEventListener("change", async (e) => {
      const files = e.target.files;
      const tableBody = document.querySelector("#dataTable tbody");
      tableBody.innerHTML = "";
      registros = [];
      duplicados = [];
  
      for (const file of files) {
        const reader = new FileReader();
        reader.onload = async function() {
          const typedArray = new Uint8Array(this.result);
          const pdf = await pdfjsLib.getDocument(typedArray).promise;
  
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const text = textContent.items.map(item => item.str).join(" ");
  
            // Extraer datos clave
            const folio = (text.match(/FOLIO\s+(\d+)/) || [])[1] || "";
            const fecha = (text.match(/FECHA:\s+(\d{2}\/\d{2}\/\d{4})/) || [])[1] || "";
            const serie = (text.match(/SERIE\s+([A-Z0-9]+)/) || [])[1] || "";
            const numero = (text.match(/NUMERO\s+(\d+)/) || [])[1] || "";
            const galones = (text.match(/GALLONS\s+([\d,\.]+)/) || [])[1] || "";
            const valorUS = (text.match(/\$([\d,]+\.\d{2})/) || [])[1] || "";
            const valorQ = (text.match(/Q([\d,]+\.\d{2})/) || [])[1] || "";
  
            let plantaCodigo = "";
            const plantaMatch = text.match(/GU\d{2}/);
            if (plantaMatch) {
              plantaCodigo = plantaMatch[0];
            }
  
            // Crear registro con todas las columnas vacías
            const registro = {
              "FOLIO": folio,
              "FECHA": fecha,
              "SERIE": serie,
              "FACTURA": numero,
              "VILLA NUEVA": "",
              "SALCAJA": "",
              "HUEHUETENANGO": "",
              "ZACAPA": "",
              "TIKAL": "",
              "COBAN": "",
              "QUICHE": "",
              "TOTONICAPAN": "",
              "ESCUINTLA": "",
              "JUTIAPA": "",
              "VALOR GTQ": valorQ,
              "VALOR EN DLL": valorUS,
              "GALONES TOTALES": ""
            };
  
            // Llenar la columna de planta con los galones
            const plantaNombre = plantas[plantaCodigo];
            if (plantaNombre && galones) {
              registro[plantaNombre] = galones;
              registro["GALONES TOTALES"] = galones;
            }
  
            registros.push(registro);
          }
  
          // Detectar duplicados (Serie + Factura)
          const combinaciones = new Set();
          registros.forEach(reg => {
            const clave = reg["SERIE"] + "-" + reg["FACTURA"];
            if (combinaciones.has(clave)) {
              duplicados.push(clave);
            } else {
              combinaciones.add(clave);
            }
          });
  
          // Mostrar registros en tabla
          tableBody.innerHTML = "";
          registros.forEach(reg => {
            const row = document.createElement("tr");
            const clave = reg["SERIE"] + "-" + reg["FACTURA"];
            if (duplicados.includes(clave)) {
              row.style.backgroundColor = "#ffcccc"; // 🔴 marcar en rojo
              row.style.color = "#b30000";
            }
  
            columnas.forEach(col => {
              const cell = document.createElement("td");
              cell.textContent = reg[col] || "";
              row.appendChild(cell);
            });
            tableBody.appendChild(row);
          });
        };
        reader.readAsArrayBuffer(file);
      }
    });
  
    // Exportar a Excel
    document.getElementById("exportExcel").addEventListener("click", () => {
      if (registros.length === 0) {
        alert("Primero carga los PDFs");
        return;
      }
  
      if (duplicados.length > 0) {
        alert("⚠️ Se detectaron facturas repetidas (Serie + Factura). Revisa las líneas marcadas en rojo antes de descargar.");
        return;
      }
  
      const worksheet = XLSX.utils.json_to_sheet(registros);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Facturas");
  
      const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
      const blob = new Blob([excelBuffer], { type: "application/octet-stream" });
      saveAs(blob, "facturas_diarias.xlsx");
    });
  });
  