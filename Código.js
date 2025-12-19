/*******************************************
 * Código.gs - Dashboard de Ventas LavaSmart
 * Zona horaria: America/Mexico_City (GMT-6)
 *******************************************/

// Constantes de configuración
var CONFIG = {
  SPREADSHEET_ID: '10_jpvm53Jn3zo0px5_wCs8Nf2YwmRpPR-CPfHC21KQs',
  SHEET_NAME: 'export',
  TIMEZONE: 'America/Mexico_City',
  DATE_FORMAT: 'yyyy-MM-dd',
  DISPLAY_DATE_FORMAT: 'dd/MM/yyyy'
};

/**
 * doGet maneja el parámetro "page":
 *   - ?page=dashboard → sirve Dashboard.html
 *   - cualquier otro caso (o sin parámetro) → sirve Index.html
 */
function doGet(e) {
  var page = (e.parameter.page || "").toLowerCase();
  var output;

  if (page === "dashboard") {
    output = HtmlService.createHtmlOutputFromFile('Dashboard');
  } else {
    output = HtmlService.createHtmlOutputFromFile('Index');
  }

  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  output.setTitle(page === "dashboard" ? 'Dashboard de Ventas | LavaSmart' : 'Consulta de Datos | LavaSmart');
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

/**
 * Formatea fecha a "DD/MM/AAAA" usando la zona horaria correcta
 */
function formatDate(dateValue) {
  if (!dateValue) return "";
  try {
    var date = new Date(dateValue);
    if (isNaN(date.getTime())) return "";
    return Utilities.formatDate(date, CONFIG.TIMEZONE, CONFIG.DISPLAY_DATE_FORMAT);
  } catch (e) {
    Logger.log('Error formateando fecha: ' + e.message);
    return "";
  }
}

/**
 * Convierte fecha a formato YYYY-MM-DD para comparaciones
 */
function getDateKey(dateValue) {
  if (!dateValue) return "";
  try {
    var date = new Date(dateValue);
    if (isNaN(date.getTime())) return "";
    return Utilities.formatDate(date, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
  } catch (e) {
    Logger.log('Error obteniendo dateKey: ' + e.message);
    return "";
  }
}

/**
 * Formatea hora a formato 12h (ej. "4:05 pm")
 */
function formatTime(dateTime) {
  if (!dateTime) return "";
  try {
    var date = new Date(dateTime);
    if (isNaN(date.getTime())) return "";
    return Utilities.formatDate(date, CONFIG.TIMEZONE, "h:mm a").toLowerCase();
  } catch (e) {
    Logger.log('Error formateando hora: ' + e.message);
    return "";
  }
}

/**
 * Formatea número como moneda MXN
 */
function formatCurrency(amount) {
  if (amount === null || amount === undefined || amount === "") return "";
  var num = parseFloat(amount);
  if (isNaN(num)) return "$0.00";
  return "$" + num.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Obtiene los datos crudos de la hoja
 */
function getRawData() {
  try {
    var sheet = SpreadsheetApp
                  .openById(CONFIG.SPREADSHEET_ID)
                  .getSheetByName(CONFIG.SHEET_NAME);
    return sheet.getRange("A3:AR").getValues();
  } catch (e) {
    Logger.log('Error obteniendo datos: ' + e.message);
    return [];
  }
}

/**
 * Verifica si una fila cumple con los filtros de fecha
 */
function isWithinDateRange(rowDate, startDate, endDate, filterType) {
  if (filterType === 'all') return true;
  
  var rowDateKey = getDateKey(rowDate);
  if (!rowDateKey) return false;
  
  if (filterType === 'day' && startDate) {
    return rowDateKey === startDate;
  }
  
  if (filterType === 'range') {
    var withinRange = true;
    if (startDate) withinRange = withinRange && (rowDateKey >= startDate);
    if (endDate) withinRange = withinRange && (rowDateKey <= endDate);
    return withinRange;
  }
  
  return true;
}

/**
 * Verifica si una fila es válida (tiene datos relevantes)
 */
function isValidRow(row) {
  if (row[3] === '❌✏️ Folio') return false;
  return row.some(function(cell) {
    return cell && cell.toString().trim() !== "";
  });
}

/**
 * Verifica si una fila coincide con el texto de búsqueda
 */
function matchesSearchText(row, filterValue) {
  if (!filterValue) return true;
  var lowerFilter = filterValue.toLowerCase();
  return (
    (row[4] && row[4].toString().toLowerCase().includes(lowerFilter)) ||
    (row[3] && row[3].toString().toLowerCase().includes(lowerFilter)) ||
    (row[5] && row[5].toString().toLowerCase().includes(lowerFilter)) ||
    (row[29] && row[29].toString().toLowerCase().includes(lowerFilter)) ||
    (row[11] && row[11].toString().toLowerCase().includes(lowerFilter)) ||
    (row[30] && row[30].toString().toLowerCase().includes(lowerFilter)) ||
    (row[37] && row[37].toString().toLowerCase().includes(lowerFilter))
  );
}

/**
 * Filtrar datos según filtros
 */
function filterData(filterValue, startDate, endDate, filterType) {
  try {
    var data = getRawData();
    if (!data.length) return [];

    return data.filter(function(row) {
      if (!isValidRow(row)) return false;
      if (!isWithinDateRange(row[4], startDate, endDate, filterType)) return false;
      if (filterType !== 'all' && !matchesSearchText(row, filterValue)) return false;
      return true;
    })
    .map(function(row) {
      return [
        row[3],                            // [0]  Folio
        formatDate(row[4]),                // [1]  Fecha
        row[5],                            // [2]  Cliente
        row[29],                           // [3]  Teléfono
        formatCurrency(row[7]),            // [4]  C. Cotizado
        formatCurrency(row[8]),            // [5]  C. S. Adicionales
        formatCurrency(row[9]),            // [6]  Propinas
        formatCurrency(row[10]),           // [7]  Diferencia
        formatCurrency(row[11]),           // [8]  Costo Total (Monto)
        formatDate(row[12]),               // [9]  Fecha de Servicio
        row[13],                           // [10] Folios Folio Físico
        row[14],                           // [11] Carpeta de Evidencia
        row[15],                           // [12] Folio de Venta
        formatCurrency(row[16]),           // [13] Costo de Servicios Base
        row[17],                           // [14] Tipo de Venta
        row[18],                           // [15] Método de Pago
        row[19],                           // [16] Estado de Pago
        row[20],                           // [17] Banco
        row[21],                           // [18] Terminación de Tarjeta
        row[22],                           // [19] Concepto / Referencia en Banco
        row[23],                           // [20] Unidad  
        row[24],                           // [21] Técnico 1
        row[25],                           // [22] Técnico 2
        row[26],                           // [23] Persona Adicional
        row[27],                           // [24] Cotización
        row[28],                           // [25] Nombre del Cliente
        row[29],                           // [26] Teléfono (duplicado)
        row[30],                           // [27] Servicio(s)
        row[31],                           // [28] Comentarios / Observaciones
        row[32],                           // [29] Servicio No.
        formatDate(row[33]),               // [30] Fecha (duplicado)
        formatTime(row[34]),               // [31] Hora
        row[35],                           // [32] Vendedor
        row[36],                           // [33] Estado Servicio
        row[37],                           // [34] Ubicación - Calle
        row[38],                           // [35] Ubicación - Colonia
        row[39],                           // [36] Ubicación - Código postal
        row[40],                           // [37] Ubicación - Cuadrante
        row[41],                           // [38] Ubicación - Unidad
        row[42],                           // [39] Ubicación - Coordenadas
        row[43]                            // [40] Referencias
      ];
    });
  } catch (e) {
    Logger.log('Error en filterData: ' + e.message);
    return [];
  }
}

/**
 * getFilteredData: devuelve JSON para el front-end
 */
function getFilteredData(filterValue, startDate, endDate, filterType) {
  try {
    var result = filterData(filterValue, startDate, endDate, filterType);
    return JSON.stringify(result);
  } catch (e) {
    Logger.log('Error en getFilteredData: ' + e.message);
    return JSON.stringify([]);
  }
}

/**
 * getKpiData: calcula KPIs según filtros
 */
function getKpiData(filterValue, startDate, endDate, filterType) {
  try {
    var data = getRawData();
    if (!data.length) {
      return { totalRecords: 0, totalRevenue: 0, avgRevenue: 0, pendingCount: 0 };
    }

    var totalRecords = 0;
    var totalRevenue = 0;
    var pendingCount = 0;

    data.forEach(function(row) {
      if (!isValidRow(row)) return;
      if (!isWithinDateRange(row[4], startDate, endDate, filterType)) return;
      if (filterType !== 'all' && !matchesSearchText(row, filterValue)) return;

      totalRecords++;
      var revenue = parseFloat(row[11]);
      if (!isNaN(revenue)) {
        totalRevenue += revenue;
      }
      if (row[36] && row[36].toString().toLowerCase() === 'pendiente') {
        pendingCount++;
      }
    });

    return {
      totalRecords: totalRecords,
      totalRevenue: totalRevenue,
      avgRevenue: totalRecords > 0 ? totalRevenue / totalRecords : 0,
      pendingCount: pendingCount
    };
  } catch (e) {
    Logger.log('Error en getKpiData: ' + e.message);
    return { totalRecords: 0, totalRevenue: 0, avgRevenue: 0, pendingCount: 0 };
  }
}

/**
 * getDailySalesData: genera arreglo para Google Charts
 */
function getDailySalesData() {
  try {
    var data = getRawData();
    var dailyMap = {};

    data.forEach(function(row) {
      if (!isValidRow(row)) return;
      var dateKey = getDateKey(row[4]);
      if (!dateKey) return;
      
      var revenue = parseFloat(row[11]);
      if (isNaN(revenue)) revenue = 0;

      if (!dailyMap[dateKey]) dailyMap[dateKey] = revenue;
      else dailyMap[dateKey] += revenue;
    });

    var sortedDates = Object.keys(dailyMap).sort();
    var result = [["Fecha", "Ventas"]];
    sortedDates.forEach(function(d) {
      result.push([d, dailyMap[d]]);
    });
    return result;
  } catch (e) {
    Logger.log('Error en getDailySalesData: ' + e.message);
    return [["Fecha", "Ventas"]];
  }
}

/**
 * getPaymentMethodData: datos para gráfica de métodos de pago
 */
function getPaymentMethodData() {
  try {
    var data = getRawData();
    var methodMap = {};

    data.forEach(function(row) {
      if (!isValidRow(row)) return;
      var method = row[18] || 'Sin especificar';
      var revenue = parseFloat(row[11]);
      if (isNaN(revenue)) revenue = 0;

      if (!methodMap[method]) methodMap[method] = revenue;
      else methodMap[method] += revenue;
    });

    var result = [["Método de Pago", "Monto"]];
    Object.keys(methodMap).forEach(function(method) {
      if (methodMap[method] > 0) {
        result.push([method, methodMap[method]]);
      }
    });
    return result;
  } catch (e) {
    Logger.log('Error en getPaymentMethodData: ' + e.message);
    return [["Método de Pago", "Monto"]];
  }
}

/**
 * getTechnicianData: datos para gráfica de técnicos
 */
function getTechnicianData() {
  try {
    var data = getRawData();
    var techMap = {};

    data.forEach(function(row) {
      if (!isValidRow(row)) return;
      var tech1 = row[24] || '';
      var tech2 = row[25] || '';
      var revenue = parseFloat(row[11]);
      if (isNaN(revenue)) revenue = 0;

      if (tech1) {
        if (!techMap[tech1]) techMap[tech1] = 0;
        techMap[tech1] += revenue;
      }
      if (tech2 && tech2 !== tech1) {
        if (!techMap[tech2]) techMap[tech2] = 0;
        techMap[tech2] += revenue;
      }
    });

    var result = [["Técnico", "Ventas"]];
    var sorted = Object.keys(techMap).sort(function(a, b) {
      return techMap[b] - techMap[a];
    });
    sorted.slice(0, 10).forEach(function(tech) {
      result.push([tech, techMap[tech]]);
    });
    return result;
  } catch (e) {
    Logger.log('Error en getTechnicianData: ' + e.message);
    return [["Técnico", "Ventas"]];
  }
}

/**
 * getVendorData: datos para gráfica de vendedores
 */
function getVendorData() {
  try {
    var data = getRawData();
    var vendorMap = {};

    data.forEach(function(row) {
      if (!isValidRow(row)) return;
      var vendor = row[35] || 'Sin asignar';
      var revenue = parseFloat(row[11]);
      if (isNaN(revenue)) revenue = 0;

      if (!vendorMap[vendor]) vendorMap[vendor] = 0;
      vendorMap[vendor] += revenue;
    });

    var result = [["Vendedor", "Ventas"]];
    var sorted = Object.keys(vendorMap).sort(function(a, b) {
      return vendorMap[b] - vendorMap[a];
    });
    sorted.forEach(function(vendor) {
      if (vendorMap[vendor] > 0) {
        result.push([vendor, vendorMap[vendor]]);
      }
    });
    return result;
  } catch (e) {
    Logger.log('Error en getVendorData: ' + e.message);
    return [["Vendedor", "Ventas"]];
  }
}

/**
 * exportToCSV: genera CSV de los datos filtrados
 */
function exportToCSV(filterValue, startDate, endDate, filterType) {
  try {
    var data = filterData(filterValue, startDate, endDate, filterType);
    if (!data.length) return "";

    var headers = [
      "Folio", "Fecha", "Cliente", "Teléfono", "C. Cotizado", "C. S. Adicionales",
      "Propinas", "Diferencia", "Costo Total", "Fecha Servicio", "Folio Físico",
      "Carpeta Evidencia", "Folio Venta", "Costo Base", "Tipo Venta", "Método Pago",
      "Estado Pago", "Banco", "Terminación Tarjeta", "Referencia", "Unidad",
      "Técnico 1", "Técnico 2", "Persona Adicional", "Cotización", "Nombre Cliente",
      "Teléfono 2", "Servicios", "Comentarios", "Servicio No", "Fecha 2", "Hora",
      "Vendedor", "Estado Servicio", "Calle", "Colonia", "CP", "Cuadrante",
      "Unidad Ubicación", "Coordenadas", "Referencias"
    ];

    var csv = headers.join(",") + "\n";
    data.forEach(function(row) {
      var csvRow = row.map(function(cell) {
        var str = (cell || "").toString().replace(/"/g, '""');
        return '"' + str + '"';
      });
      csv += csvRow.join(",") + "\n";
    });

    return csv;
  } catch (e) {
    Logger.log('Error en exportToCSV: ' + e.message);
    return "";
  }
}

/**
 * getServiceStatusData: datos para gráfica de estados de servicio
 */
function getServiceStatusData() {
  try {
    var data = getRawData();
    var statusMap = {};

    data.forEach(function(row) {
      if (!isValidRow(row)) return;
      var status = row[36] || 'Sin estado';
      if (!statusMap[status]) statusMap[status] = 0;
      statusMap[status]++;
    });

    var result = [["Estado", "Cantidad"]];
    Object.keys(statusMap).forEach(function(status) {
      result.push([status, statusMap[status]]);
    });
    return result;
  } catch (e) {
    Logger.log('Error en getServiceStatusData: ' + e.message);
    return [["Estado", "Cantidad"]];
  }
}
