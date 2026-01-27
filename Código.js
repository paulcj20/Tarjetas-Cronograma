// ============================================================================
// CONFIGURACIÓN: PEGA AQUÍ LA URL DE TU ENDPOINT DE GOOGLE APPS SCRIPT
// ============================================================================
// Para crear el endpoint:
// 1. En tu hoja de Google Sheets, ve a Extensiones > Apps Script
// 2. Crea una función doGet() que devuelva los datos en JSON (ver ejemplo abajo)
// 3. Publica como aplicación web (Implementar > Nueva implementación > Aplicación web)
// 4. Selecciona "Cualquier persona" en "Quién tiene acceso"
// 5. Copia la URL que te da y pégala aquí:

const URL_API = 'https://script.google.com/macros/s/AKfycbyTeVfgta4PC27Xzg1b5BANqap9WsYT1jTbeyLYuUgBkrkkgLnGCdnBDcGJGmXaGYBNaA/exec';

// Ejemplo de función doGet() para Apps Script:
/*
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('G. Nov 25'); // Ajusta el nombre
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  const json = rows.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
  
  return ContentService.createTextOutput(JSON.stringify(json))
    .setMimeType(ContentService.MimeType.JSON);
}
*/
  // ============================================================================
// VARIABLES GLOBALES
// ============================================================================
let allData = []; // Datos del mes seleccionado
let filteredData = []; // Datos filtrados del mes seleccionado
let filteredHistoricalData = []; // Datos filtrados acumulados por mes
const historicalData = {}; // Datos normalizados por planilla
let currentMonth = '';
let chartMargenMensual = null;
let chartMargenCliente = null;
let chartMarginWaterfall = null;
let chartCorredorTarifa = null;
// Ajusta este valor para cambiar la tarifa por hora del chofer en el Waterfall.
const WATERFALL_DEFAULT_COSTO_HORA = 12;
  const MONTH_NAME_MAP = {
    ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6,
    jul: 7, ago: 8, sep: 9, oct: 10, nov: 11, dic: 12
};
  // Convierte valores con separadores locales a números válidos
function toNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    if (value === null || value === undefined) return 0;
    const cleaned = value
        .toString()
        .trim()
        .replace(/\s+/g, '')
        .replace(/\./g, '')
        .replace(',', '.');
    const parsed = parseFloat(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
}
  // Normaliza encabezados para hacer coincidencias robustas
function normalizeKey(key) {
    if (key === null || key === undefined) return '';
    return key
        .toString()
        .trim()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]+/g, '');
}
  function buildRowMap(row) {
    const map = {};
    if (!row || typeof row !== 'object') return map;
    Object.keys(row).forEach(key => {
        const normalized = normalizeKey(key);
        if (normalized && map[normalized] === undefined) {
            map[normalized] = row[key];
        }
    });
    return map;
}
  function getValue(map, header, fallback = '') {
    const targets = Array.isArray(header) ? header : [header];
    for (const target of targets) {
        const normalized = normalizeKey(target);
        if (!normalized) continue;
        if (Object.prototype.hasOwnProperty.call(map, normalized)) {
            const value = map[normalized];
            if (value !== undefined && value !== null && value !== '') return value;
        }
    }
    return fallback;
}
  function getNumber(map, header) {
    return toNumber(getValue(map, header, 0));
}
  function normalizeText(value) {
    if (value === null || value === undefined) return '';
    return value.toString().trim().toLowerCase();
}
  // ============================================================================
// FUNCIÓN: Cargar datos de la API usando JSONP (evita problemas de CORS)
// ============================================================================
function fetchData() {
    const selectMes = document.getElementById('filterMes');
    const mesSeleccionado = selectMes ? selectMes.value : '';
    currentMonth = mesSeleccionado;
      console.log('🔄 fetchData: Iniciando carga de datos para', mesSeleccionado || '(Todos los meses)');
    console.log('📍 URL_API:', URL_API);
      if (!mesSeleccionado) {
        console.log('📅 "Todos" seleccionado - cargando todos los meses disponibles');
        const availableMonths = getAvailableMonths();
        
        if (availableMonths.length === 0) {
            console.warn('⚠️ No hay meses disponibles en el dropdown.');
            return;
        }
        
        // Precargar todos los meses y luego mostrar datos consolidados
        let loadedCount = 0;
        availableMonths.forEach(month => {
            if (!historicalData[month]) {
                requestSheetData(
                    month,
                    data => {
                        if (Array.isArray(data)) {
                            historicalData[month] = normalizeDataset(data, month);
                            console.log(`📦 Mes cargado (${month}):`, historicalData[month].length, 'registros');
                        } else {
                            console.warn(`⚠️ Datos inválidos para ${month}`);
                            historicalData[month] = [];
                        }
                        loadedCount++;
                        
                        // Cuando todos los meses estén cargados, procesar datos consolidados
                        if (loadedCount === availableMonths.length) {
                            processDataAllMonths();
                        }
                    },
                    () => {
                        console.warn(`⚠️ Error al cargar ${month}`);
                        loadedCount++;
                        if (loadedCount === availableMonths.length) {
                            processDataAllMonths();
                        }
                    },
                    { silent: true }
                );
            }
        });
        
        // Si ya todos los meses están cargados, procesar inmediatamente
        if (availableMonths.every(month => historicalData[month])) {
            processDataAllMonths();
        }
        return;
    }
      requestSheetData(
        mesSeleccionado,
        data => processData(data, mesSeleccionado),
        () => {
            document.getElementById('tableWrapper').innerHTML =
                '<div class="error">Error al cargar datos. Verifica la URL del endpoint.<br>URL: ' + URL_API + '</div>';
        }
    );
}
  function requestSheetData(sheetName, onSuccess, onError, options = {}) {
    const { silent = false } = options;
    const callbackName = `jsonpCallback_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
    const finalURL = `${URL_API}?callback=${callbackName}&sheet=${encodeURIComponent(sheetName)}`;
    const script = document.createElement('script');
      console.log(`📡 requestSheetData: ${sheetName} → ${finalURL}`);
      window[callbackName] = function(data) {
        console.log(`✅ Datos recibidos (${sheetName}):`, data.length, 'filas');
        cleanup();
        if (onSuccess) onSuccess(data, sheetName);
    };
      function cleanup() {
        delete window[callbackName];
        if (script.parentNode) {
            script.parentNode.removeChild(script);
        }
    }
      script.src = finalURL;
    script.onerror = function(error) {
        console.error(`❌ Error al cargar datos (${sheetName}):`, error);
        cleanup();
        if (!silent) {
            console.error('❌ Error crítico al cargar planilla principal.');
        }
        if (onError) onError(error, sheetName);
    };
      script.onload = function() {
        console.log(`📥 Script cargado (${sheetName})`);
    };
      document.body.appendChild(script);
}
  function getAvailableMonths() {
    const select = document.getElementById('filterMes');
    if (!select) return [];
    return Array.from(select.options)
        .map(option => option.value)
        .filter(Boolean);
}
  function normalizeMonthToken(name) {
    return normalizeKey(name).slice(0, 3);
}
  function getMonthKeyFromSheet(sheet) {
    if (!sheet) return '';
    const parts = sheet.toString().trim().split(/\s+/);
    if (parts.length < 2) return '';
    const monthToken = normalizeMonthToken(parts[0]);
    const monthNumber = MONTH_NAME_MAP[monthToken];
    const yearRaw = parts[parts.length - 1];
    const yearParsed = parseInt(yearRaw, 10);
    if (!monthNumber || isNaN(yearParsed)) return '';
    const fullYear = yearParsed < 100 ? 2000 + yearParsed : yearParsed;
    return `${fullYear}-${String(monthNumber).padStart(2, '0')}`;
}
  function monthKeyToLabel(key) {
    if (!key) return '';
    const [yearStr, monthStr] = key.split('-');
    const year = parseInt(yearStr, 10);
    const monthIndex = parseInt(monthStr, 10);
    if (!year || !monthIndex) return key;
    const monthNames = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
    return `${monthNames[monthIndex - 1]} ${year}`;
}
  function getHistoricalRows() {
    return Object.values(historicalData).flat();
}
  function preloadAdditionalMonths(selectedMonth) {
    const months = getAvailableMonths()
        .filter(month => month && month !== selectedMonth && !historicalData[month]);
      if (!months.length) return;
      months.forEach(month => {
        requestSheetData(
            month,
            data => {
                historicalData[month] = normalizeDataset(data, month);
                console.log(`📦 Planilla precargada (${month}):`, historicalData[month].length, 'registros');
                const filters = getActiveFilters();
                filteredHistoricalData = filterDataset(getHistoricalRows(), filters);
                renderCharts();
            },
            null,
            { silent: true }
        );
    });
}
  function normalizeDataset(data, monthKey) {
    if (!Array.isArray(data)) {
        console.warn(`⚠️ normalizeDataset: data no es un array para ${monthKey}`, typeof data, data);
        return [];
    }
    return data.map(row => {
        const map = buildRowMap(row);
        const kilometros = getNumber(map, ['Kilometros', 'KM Promedio', 'KM']);
        const record = {
            planillaMes: monthKey,
            carpeta: getValue(map, 'Carpeta'),
            factura: getValue(map, 'Factura'),
            fecha: getValue(map, ['Fecha de OP.', 'Fecha']),
            contenedor: getValue(map, ['Nro. Cont - OP.', 'Contenedor']),
            tipoOP: getValue(map, 'Tipo de OP.'),
            matricula: getValue(map, 'Matricula'),
            cliente: getValue(map, 'Cliente'),
            origen: getValue(map, 'Origen'),
            destino: getValue(map, 'Destino'),
            terciarizado: getValue(map, 'Terciarizado', 'No') || 'No',
            ventaFlete: getNumber(map, 'Venta Flete'),
            costosExtra: getNumber(map, 'Costos Extra'),
            detExtras: getValue(map, ['Det. Extras', 'Det Extras']),
            costoFletero: getNumber(map, 'Costo Fletero'),
            extrProveedores: getNumber(map, ['Extr. Proveedores', 'Costos OP.']),
            costoChofer: getNumber(map, 'Costo Chofer'),
            horaInicio: getValue(map, ['Hora inicio', 'Inicio Op.', 'Inicio']),
            horaFinal: getValue(map, ['Hora final', 'Fin Op.', 'Fin']),
            kilometros,
            litros: getNumber(map, 'Litros'),
            costoGasoil: getNumber(map, 'Costo Gasoil'),
            totalCostos: getNumber(map, 'Total de Costos'),
            totalVenta: getNumber(map, 'Total Venta'),
            get margen() {
                return (this.ventaFlete + this.costosExtra) - this.totalCostos;
            },
            get margenPct() {
                const ingresos = this.ventaFlete + this.costosExtra;
                return ingresos > 0 ? (this.margen / ingresos * 100) : 0;
            }
        };
          record.horas = computeHoras(record);
        return record;
    });
}
  function parseTime(value) {
    if (!value) return null;
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    const str = value.toString().trim();
    if (!str) return null;
      const timeMatch = str.match(/^\s*(\d{1,2})(?::(\d{2}))?\s*$/);
    if (timeMatch) {
        const hours = parseInt(timeMatch[1], 10);
        const minutes = timeMatch[2] ? parseInt(timeMatch[2], 10) : 0;
        if (isNaN(hours) || hours < 0 || hours > 23) return null;
        if (isNaN(minutes) || minutes < 0 || minutes > 59) return null;
        const date = new Date();
        date.setHours(hours, minutes, 0, 0);
        return date;
    }
      const date = new Date(str);
    return isNaN(date.getTime()) ? null : date;
}
  function computeHoras(row) {
    if (!row) return '';
      const terciarizadoKey = normalizeKey(row.terciarizado);
    if (terciarizadoKey === 'si') return 'N/A';
      const inicio = parseTime(row.horaInicio);
    const fin = parseTime(row.horaFinal);
    if (!inicio || !fin) return '';
      let diffMs = fin.getTime() - inicio.getTime();
    if (diffMs < 0) {
        diffMs += 24 * 60 * 60 * 1000;
    }
      const hours = diffMs / (60 * 60 * 1000);
    return hours.toFixed(1);
}
  function processData(data, monthKey) {
    try {
        console.log('🔧 processData: Procesando', data.length, 'registros para', monthKey);
        
        // Debug: Ver los campos disponibles
        if (data.length > 0) {
            console.log('📋 Campos disponibles:', Object.keys(data[0]));
            console.log('📋 Primera fila completa:', data[0]);
        }
        
        const normalized = normalizeDataset(data, monthKey);
        historicalData[monthKey] = normalized;
        allData = normalized;
          console.log('📊 Datos normalizados:', allData.length, 'registros');
        console.log('🚗 Primeros 3 registros con origen/km:', allData.slice(0, 3).map(r => ({
            carpeta: r.carpeta,
            origen: r.origen,
            km: r.kilometros
        })));
        
        // Inicializar filtros con opciones únicas
        populateFilters();
        
        // Aplicar filtros iniciales (todos los datos)
        applyFilters();
          // Precargar otras planillas para la gráfica mensual
        preloadAdditionalMonths(monthKey);
        
        console.log('✅ Dashboard cargado exitosamente para', monthKey);
        
    } catch (error) {
        console.error('❌ Error en processData:', error);
        document.getElementById('tableWrapper').innerHTML = 
            `<div class="error">Error al procesar datos: ${error.message}</div>`;
    }
}
  function processDataAllMonths() {
    try {
        console.log('🔧 processDataAllMonths: Procesando todos los meses consolidados');
        
        // Consolidar todos los datos de los meses cargados
        allData = getHistoricalRows();
        
        console.log('📊 Datos consolidados:', allData.length, 'registros totales');
        
        // Inicializar filtros con opciones únicas
        populateFilters();
        
        // Aplicar filtros iniciales (todos los datos)
        applyFilters();
        
        console.log('✅ Dashboard cargado exitosamente (Todos los meses)');
        
    } catch (error) {
        console.error('❌ Error en processDataAllMonths:', error);
        document.getElementById('tableWrapper').innerHTML = 
            `<div class="error">Error al procesar datos consolidados: ${error.message}</div>`;
    }
}
  // ============================================================================
// FUNCIÓN: Poblar selectores de filtros con opciones únicas
// ============================================================================
function populateFilters() {
    // Limpiar selectores existentes
    const selectCliente = document.getElementById('filterCliente');
    selectCliente.innerHTML = '<option value="">Todos</option>';
    
    const selectTipoOP = document.getElementById('filterTipoOP');
    selectTipoOP.innerHTML = '<option value="">Todos</option>';
    
    const selectTerciarizado = document.getElementById('filterTerciarizado');
    selectTerciarizado.value = '';
    
    // Clientes únicos
    const clientes = [...new Set(allData.map(d => d.cliente))].filter(c => c).sort();
    clientes.forEach(cliente => {
        const option = document.createElement('option');
        option.value = cliente;
        option.textContent = cliente;
        selectCliente.appendChild(option);
    });
      // Tipos de OP únicos
    const tiposOP = [...new Set(allData.map(d => d.tipoOP))].filter(t => t).sort();
    tiposOP.forEach(tipo => {
        const option = document.createElement('option');
        option.value = tipo;
        option.textContent = tipo;
        selectTipoOP.appendChild(option);
    });
}
  function getActiveFilters() {
    return {
        cliente: document.getElementById('filterCliente').value,
        tipoOP: document.getElementById('filterTipoOP').value,
        terciarizado: document.getElementById('filterTerciarizado').value,
        searchCarpeta: normalizeText(document.getElementById('searchCarpeta')?.value),
        searchContenedor: normalizeText(document.getElementById('searchContenedor')?.value)
    };
}
  function matchesFilters(row, filters) {
    if (filters.cliente && row.cliente !== filters.cliente) return false;
    if (filters.tipoOP && row.tipoOP !== filters.tipoOP) return false;
    if (filters.terciarizado && row.terciarizado !== filters.terciarizado) return false;
    if (filters.searchCarpeta && !normalizeText(row.carpeta).includes(filters.searchCarpeta)) return false;
    if (filters.searchContenedor && !normalizeText(row.contenedor).includes(filters.searchContenedor)) return false;
    return true;
}
  function filterDataset(dataset, filters) {
    return dataset
        .filter(row => matchesFilters(row, filters));
}
  // ============================================================================
// FUNCIÓN: Aplicar filtros a los datos
// ============================================================================
function applyFilters() {
    const filters = getActiveFilters();
    filteredData = filterDataset(allData, filters);
    filteredHistoricalData = filterDataset(getHistoricalRows(), filters);
      // Actualizar toda la interfaz
    calculateKPIs();
    renderCharts();
    renderTable();
}
  // ============================================================================
// FUNCIÓN: Calcular KPIs en base a datos filtrados
// ============================================================================
function calculateKPIs() {
    const ingresoTotal = filteredData.reduce((sum, row) => sum + row.totalVenta, 0);
    const costoTotal = filteredData.reduce((sum, row) => sum + row.totalCostos, 0);
    const margenBruto = ingresoTotal - costoTotal;
    const margenPct = ingresoTotal > 0 ? (margenBruto / ingresoTotal * 100) : 0;
    
    // Solo contar KM de operaciones que realmente tienen KM > 0
    const totalKm = filteredData.reduce((sum, row) => {
        return sum + (row.kilometros > 0 ? row.kilometros : 0);
    }, 0);
    const costoKm = totalKm > 0 ? (costoTotal / totalKm) : 0;
    
    // Debug detallado
    console.log('💰 KPIs:', {
        registros: filteredData.length,
        costoTotal: costoTotal,
        totalKm: totalKm,
        costoKm: costoKm,
        ejemploKm: filteredData.slice(0, 5).map(r => r.kilometros),
        registrosConKm: filteredData.filter(r => r.kilometros > 0).length
    });
      // Actualizar UI
    document.getElementById('kpiIngreso').textContent = formatCurrency(ingresoTotal);
    document.getElementById('kpiCosto').textContent = formatCurrency(costoTotal);
    
    const margenEl = document.getElementById('kpiMargen');
    margenEl.textContent = formatCurrency(margenBruto);
    margenEl.className = 'kpi-value ' + (margenBruto >= 0 ? 'positive' : 'negative');
    
    const margenPctEl = document.getElementById('kpiMargenPct');
    margenPctEl.textContent = margenPct.toFixed(1) + '%';
    margenPctEl.className = 'kpi-value ' + (margenPct >= 0 ? 'positive' : 'negative');
    
    document.getElementById('kpiCostoKm').textContent = formatCurrency(costoKm, 2);
}
  // ============================================================================
// FUNCIÓN: Renderizar gráficos con Chart.js
// ============================================================================
function renderCharts() {
    renderMargenMensualChart(filteredHistoricalData);
    renderMargenClienteChart();
  renderMarginWaterfallChart();
    renderTarifaVsObjetivoPorCorredor(filteredData);
}
  // Gráfico de líneas: Evolución mensual del margen bruto
function renderMargenMensualChart(dataset) {
    const source = Array.isArray(dataset) && dataset.length ? dataset : filteredData;
    const margenPorMes = {};
    
    source.forEach(row => {
        const referenceMonth = row.planillaMes || currentMonth;
        let mesKey = getMonthKeyFromSheet(referenceMonth);
          if (!mesKey) {
            const date = parseDate(row.fecha, referenceMonth);
            if (date) {
                mesKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            }
        }
          if (!mesKey) return;
          if (!margenPorMes[mesKey]) {
            margenPorMes[mesKey] = { ingreso: 0, costo: 0 };
        }
          margenPorMes[mesKey].ingreso += row.totalVenta;
        margenPorMes[mesKey].costo += row.totalCostos;
    });
      const meses = Object.keys(margenPorMes).sort();
    const margenes = meses.map(mes => margenPorMes[mes].ingreso - margenPorMes[mes].costo);
    const labels = meses.map(monthKeyToLabel);
      // Destruir gráfico anterior si existe
    if (chartMargenMensual) {
        chartMargenMensual.destroy();
    }
      const ctx = document.getElementById('chartMargenMensual').getContext('2d');
    chartMargenMensual = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Margen Bruto (USD)',
                data: margenes,
                borderColor: '#3b82f6',
                backgroundColor: 'rgba(59, 130, 246, 0.1)',
                tension: 0.4,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return '$' + value.toLocaleString();
                        }
                    }
                }
            }
        }
    });
}
  // Gráfico de barras: Margen % por cliente
function renderMargenClienteChart() {
    // Agrupar datos por cliente
    const margenPorCliente = {};
    
    filteredData.forEach(row => {
        if (!row.cliente) return;
        
        if (!margenPorCliente[row.cliente]) {
            margenPorCliente[row.cliente] = { ingreso: 0, costo: 0 };
        }
        
        margenPorCliente[row.cliente].ingreso += row.totalVenta;
        margenPorCliente[row.cliente].costo += row.totalCostos;
    });
      // Calcular margen % por cliente y ordenar
    const clientes = Object.keys(margenPorCliente).map(cliente => ({
        nombre: cliente,
        margenPct: margenPorCliente[cliente].ingreso > 0 
            ? ((margenPorCliente[cliente].ingreso - margenPorCliente[cliente].costo) / margenPorCliente[cliente].ingreso * 100)
            : 0
    })).sort((a, b) => b.margenPct - a.margenPct).slice(0, 10); // Top 10
      // Destruir gráfico anterior si existe
    if (chartMargenCliente) {
        chartMargenCliente.destroy();
    }
      const ctx = document.getElementById('chartMargenCliente').getContext('2d');
    chartMargenCliente = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: clientes.map(c => c.nombre),
            datasets: [{
                label: 'Margen %',
                data: clientes.map(c => c.margenPct),
                backgroundColor: clientes.map(c => c.margenPct >= 0 ? '#10b981' : '#ef4444')
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                }
            }
        }
    });
}

const waterfallValueLabelsPlugin = {
    id: 'waterfallValueLabels',
    afterDatasetsDraw(chart) {
        const { ctx } = chart;
        const datasets = chart.data.datasets || [];
        const fontFamily = chart.options.font?.family || 'Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI"';
        const rootStyles = getComputedStyle(document.documentElement);
        const fallbackColor = '#1f2937';
        const textColor = chart.options.plugins?.waterfallLabels?.color?.trim() ||
            rootStyles.getPropertyValue('--color-text-primary').trim() ||
            fallbackColor;

        ctx.save();
        ctx.font = `12px ${fontFamily}`;
        ctx.fillStyle = textColor;
        ctx.textAlign = 'center';

        datasets.forEach((dataset, datasetIndex) => {
            const meta = chart.getDatasetMeta(datasetIndex);
            if (!meta) return;

            dataset.data?.forEach((dataPoint, index) => {
                if (!dataPoint) return;
                const element = meta.data?.[index];
                if (!element || element.hidden || element.skip) return;
                const text = dataPoint.labelText;
                if (!text) return;

                const position = element.tooltipPosition();
                const offset = dataPoint.delta >= 0 ? -8 : 16;
                ctx.fillText(text, position.x, position.y + offset);
            });
        });

        ctx.restore();
    }
};

function buildMarginWaterfall(dataset, options = {}) {
    // Para agregar nuevas categorías (peajes, mantenimiento, etc.), extiende costSteps más abajo.
    const settings = {
        costoHoraChofer: options.costoHoraChofer ?? WATERFALL_DEFAULT_COSTO_HORA,
        currency: options.currency || 'USD', // Cambia la moneda aquí.
        locale: options.locale || 'es-UY',
        currencyDigits: options.currencyDigits ?? 0,
        dateRange: options.dateRange || null
    };

    const formatter = new Intl.NumberFormat(settings.locale, {
        style: 'currency',
        currency: settings.currency,
        minimumFractionDigits: settings.currencyDigits,
        maximumFractionDigits: settings.currencyDigits
    });

    let startDate = null;
    let endDate = null;
    if (settings.dateRange) {
        if (settings.dateRange.start) {
            startDate = settings.dateRange.start instanceof Date
                ? settings.dateRange.start
                : new Date(settings.dateRange.start);
            if (Number.isNaN(startDate?.getTime())) startDate = null;
        }
        if (settings.dateRange.end) {
            endDate = settings.dateRange.end instanceof Date
                ? settings.dateRange.end
                : new Date(settings.dateRange.end);
            if (Number.isNaN(endDate?.getTime())) endDate = null;
        }
    }

    const rows = Array.isArray(dataset) ? dataset.filter(row => {
        if (!settings.dateRange) return true;
        const parsed = parseDate(row.fecha, row.planillaMes || currentMonth);
        if (!parsed) return false;
        if (startDate && parsed < startDate) return false;
        if (endDate && parsed > endDate) return false;
        return true;
    }) : [];

    function extractNumber(source, keys) {
        for (const key of keys) {
            if (source[key] !== undefined && source[key] !== null && source[key] !== '') {
                const value = toNumber(source[key]);
                if (Number.isFinite(value)) {
                    return value;
                }
            }
        }
        return null;
    }

    function numberOr(value, fallback = 0) {
        if (value === null || value === undefined || Number.isNaN(value)) {
            return typeof fallback === 'number' ? fallback : toNumber(fallback);
        }
        return value;
    }

    function sanitizeHours(value) {
        if (typeof value === 'number') {
            return Number.isFinite(value) ? value : 0;
        }
        if (!value) return 0;
        const parsed = parseFloat(value.toString().replace(',', '.'));
        return Number.isFinite(parsed) ? parsed : 0;
    }

    const totals = rows.reduce((acc, row) => {
        const ingresoTotal = numberOr(extractNumber(row, ['ingreso', 'totalVenta', 'ventaFlete']), 0);
        acc.ingresos += ingresoTotal;

        const costoTercerosDirect = extractNumber(row, ['costoTerceros']);
        const costoFletero = numberOr(extractNumber(row, ['costoFletero']), 0);
        const costoProveedores = numberOr(extractNumber(row, ['extrProveedores']), 0);
        const costoTerceros = costoTercerosDirect !== null ? costoTercerosDirect : (costoFletero + costoProveedores);
        acc.costoTerceros += costoTerceros;

        const combustible = numberOr(extractNumber(row, ['combustible', 'costoGasoil']), 0);
        acc.combustible += combustible;

        const horas = sanitizeHours(row.horasChofer ?? row.horas);
        const tarifa = numberOr(extractNumber(row, ['costoHoraChofer']), settings.costoHoraChofer);
        let costoChofer = horas * tarifa;
        if (!horas) {
            costoChofer = numberOr(extractNumber(row, ['costoChofer']), costoChofer);
        }
        acc.costoChofer += costoChofer;
        acc.horasTotales += horas;

        return acc;
    }, { ingresos: 0, costoTerceros: 0, combustible: 0, costoChofer: 0, horasTotales: 0 });

    const margin = totals.ingresos - totals.costoTerceros - totals.combustible - totals.costoChofer;
    const marginPct = totals.ingresos > 0 ? (margin / totals.ingresos) * 100 : 0;

    let cumulative = 0;
    const steps = [];

    steps.push({
        label: 'Ingresos',
        type: 'ingreso',
        delta: totals.ingresos,
        start: 0,
        end: totals.ingresos,
        percent: totals.ingresos > 0 ? 100 : 0
    });
    cumulative = totals.ingresos;

    const costSteps = [
        { label: 'Terceros', value: totals.costoTerceros },
        { label: 'Combustible', value: totals.combustible },
        { label: 'Chofer', value: totals.costoChofer }
    ];

    costSteps.forEach(step => {
        const delta = -step.value;
        const start = cumulative;
        const end = cumulative + delta;
        cumulative = end;
        steps.push({
            label: step.label,
            type: 'costo',
            delta,
            start,
            end,
            percent: totals.ingresos > 0 ? (delta / totals.ingresos) * 100 : 0
        });
    });

    steps.push({
        label: 'Margen Bruto',
        type: 'margen',
        delta: margin,
        start: 0,
        end: margin,
        percent: totals.ingresos > 0 ? marginPct : 0
    });

    const labels = steps.map(step => step.label);

    function toDataArray(type) {
        return steps
            .filter(step => step.type === type)
            .map(step => {
                const rangeMin = Math.min(step.start, step.end);
                const rangeMax = Math.max(step.start, step.end);
                const percentValue = totals.ingresos > 0 ? step.percent : 0;
                const percentText = totals.ingresos > 0 ? ` (${percentValue.toFixed(1)}%)` : '';
                return {
                    x: step.label,
                    y: [rangeMin, rangeMax],
                    delta: step.delta,
                    percent: percentValue,
                    cumulative: step.end,
                    labelText: `${formatter.format(step.delta)}${percentText}`
                };
            });
    }

    const axisValues = steps.flatMap(step => [step.start, step.end]);
    let minAxis = Math.min(0, ...axisValues);
    let maxAxis = Math.max(0, ...axisValues);
    if (minAxis === maxAxis) {
        minAxis -= 1;
        maxAxis += 1;
    }

    return {
        labels,
        datasets: {
            ingresos: toDataArray('ingreso'),
            costos: toDataArray('costo'),
            margen: toDataArray('margen')
        },
        formatter,
        totals,
        marginPct,
        axis: {
            min: minAxis - Math.abs(minAxis) * 0.08,
            max: maxAxis + Math.abs(maxAxis) * 0.08
        },
        settings
    };
}

function renderMarginWaterfallChart() {
    const canvas = document.getElementById('chartMarginWaterfall');
    if (!canvas) return;

    const wrapper = canvas.parentElement;
    if (wrapper) {
        wrapper.style.height = '320px';
    }

    const chartPayload = buildMarginWaterfall(filteredData, {
        costoHoraChofer: WATERFALL_DEFAULT_COSTO_HORA,
        currency: 'USD',
        locale: 'es-UY',
        currencyDigits: 0
    });

    if (chartMarginWaterfall) {
        chartMarginWaterfall.destroy();
    }

    const styles = getComputedStyle(document.documentElement);
    const incomeColor = (styles.getPropertyValue('--color-positive') || '#10b981').trim() || '#10b981';
    const costColor = (styles.getPropertyValue('--color-negative') || '#ef4444').trim() || '#ef4444';
    const marginColor = (styles.getPropertyValue('--color-accent-strong') || '#1d4ed8').trim() || '#1d4ed8';
    const textColor = (styles.getPropertyValue('--color-text-primary') || '#1f2937').trim() || '#1f2937';
    const surfaceColor = (styles.getPropertyValue('--color-surface') || '#ffffff').trim() || '#ffffff';
    const gridColor = (styles.getPropertyValue('--color-border') || 'rgba(148, 163, 184, 0.3)').trim() || 'rgba(148, 163, 184, 0.3)';

    chartMarginWaterfall = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels: chartPayload.labels,
            datasets: [
                {
                    label: 'Ingresos',
                    data: chartPayload.datasets.ingresos,
                    backgroundColor: incomeColor,
                    borderColor: incomeColor,
                    borderWidth: 0,
                    borderRadius: 6,
                    borderSkipped: false,
                    hoverBackgroundColor: incomeColor
                },
                {
                    label: 'Costos',
                    data: chartPayload.datasets.costos,
                    backgroundColor: costColor,
                    borderColor: costColor,
                    borderWidth: 0,
                    borderRadius: 6,
                    borderSkipped: false,
                    hoverBackgroundColor: costColor
                },
                {
                    label: 'Margen',
                    data: chartPayload.datasets.margen,
                    backgroundColor: marginColor,
                    borderColor: marginColor,
                    borderWidth: 0,
                    borderRadius: 6,
                    borderSkipped: false,
                    hoverBackgroundColor: marginColor
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            parsing: {
                xAxisKey: 'x',
                yAxisKey: 'y'
            },
            animation: {
                duration: 600,
                easing: 'easeOutCubic'
            },
            plugins: {
                legend: {
                    labels: {
                        usePointStyle: true,
                        color: textColor
                    }
                },
                tooltip: {
                    backgroundColor: surfaceColor,
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: gridColor,
                    borderWidth: 1,
                    callbacks: {
                        label(context) {
                            const raw = context.raw;
                            if (!raw) return '';
                            const valueText = chartPayload.formatter.format(raw.delta);
                            if (chartPayload.totals.ingresos > 0) {
                                return `${context.dataset.label}: ${valueText} (${raw.percent.toFixed(1)}%)`;
                            }
                            return `${context.dataset.label}: ${valueText}`;
                        },
                        footer() {
                            if (chartPayload.totals.ingresos <= 0) return '';
                            return `Margen Bruto %: ${chartPayload.marginPct.toFixed(1)}%`;
                        }
                    }
                },
                waterfallLabels: {
                    color: textColor
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: textColor
                    },
                    grid: {
                        display: false
                    }
                },
                y: {
                    ticks: {
                        color: textColor,
                        callback(value) {
                            return chartPayload.formatter.format(value);
                        }
                    },
                    grid: {
                        color: gridColor
                    },
                    min: chartPayload.axis.min,
                    max: chartPayload.axis.max
                }
            }
        },
        plugins: [waterfallValueLabelsPlugin]
    });
}

function computeMedian(values) {
    if (!Array.isArray(values) || values.length === 0) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    if (sorted.length % 2 !== 0) {
        return sorted[mid];
    }
    return (sorted[mid - 1] + sorted[mid]) / 2;
}

function prepareCorredorTarifaDataset(rows) {
    const corredorMap = new Map();

    rows.forEach(row => {
        const isTerciarizado = normalizeKey(row.terciarizado) === 'si';
        if (!isTerciarizado) return;

        const origen = typeof row.origen === 'string' ? row.origen.trim() : (row.origen || '');
        const destino = typeof row.destino === 'string' ? row.destino.trim() : (row.destino || '');
        if (!origen || !destino) return;

        const corredor = `${origen} → ${destino}`;
        if (!corredorMap.has(corredor)) {
            corredorMap.set(corredor, {
                corredor,
                viajes: 0,
                tarifaCotizadaSum: 0,
                tarifaCotizadaCount: 0,
                tarifaCotizadaTotal: 0,
                ventaTotal: 0,
                ventaSum: 0,
                ventaCount: 0,
                costValues: [],
                costPositiveCount: 0
            });
        }

        const group = corredorMap.get(corredor);
        group.viajes += 1;

        const venta = Number.isFinite(row.ventaFlete) ? row.ventaFlete : toNumber(row.ventaFlete);
        if (Number.isFinite(venta)) {
            group.ventaSum += venta;
            group.ventaTotal += venta;
            group.ventaCount += 1;
        }

        const tarifaCotizada = Number.isFinite(row.costoFletero) ? row.costoFletero : toNumber(row.costoFletero);
        if (Number.isFinite(tarifaCotizada)) {
            group.tarifaCotizadaSum += tarifaCotizada;
            group.tarifaCotizadaTotal += tarifaCotizada;
            group.tarifaCotizadaCount += 1;
        }

        const costo = Number.isFinite(row.totalCostos) ? row.totalCostos : toNumber(row.totalCostos);
        if (Number.isFinite(costo)) {
            group.costValues.push(costo);
            if (costo > 0) {
                group.costPositiveCount += 1;
            }
        }
    });

    return Array.from(corredorMap.values()).map(group => {
        const ventaPromedio = group.ventaCount > 0 ? (group.ventaSum / group.ventaCount) : 0;
        const tarifaCotizadaProm = group.tarifaCotizadaCount > 0
            ? (group.tarifaCotizadaSum / group.tarifaCotizadaCount)
            : 0;

        const medianCost = computeMedian(group.costValues);
        const averageCost = group.costValues.length
            ? group.costValues.reduce((sum, value) => sum + value, 0) / group.costValues.length
            : 0;
        const costoTipico = medianCost > 0 ? medianCost : averageCost;

        const tarifaDeseada = ventaPromedio * 0.8;
        const gap = tarifaCotizadaProm - tarifaDeseada;
        const gapPct = tarifaDeseada !== 0 ? (gap / tarifaDeseada) : 0;
        const costCompleteness = group.viajes > 0 ? (group.costPositiveCount / group.viajes) : 0;

        return {
            corredor: group.corredor,
            viajes: group.viajes,
            tarifaCotizadaProm,
            tarifaCotizadaTotal: group.tarifaCotizadaTotal,
            ventaPromedio,
            ventaTotal: group.ventaTotal,
            costoTipico,
            tarifaDeseada,
            gap,
            gapPct,
            costCompleteness,
            costDataIncomplete: costCompleteness < 0.5
        };
    }).filter(item => item.viajes > 0);
}

// Para extender el análisis con nuevos rubros (peajes, mantenimiento, etc.), ajusta
// prepareCorredorTarifaDataset para considerar los campos adicionales y recalcular la tarifa objetivo.
function renderTarifaVsObjetivoPorCorredor(rows = [], options = {}) {
    const canvas = document.getElementById('chartTarifaCorredor');
    if (!canvas) return;

    const wrapper = canvas.parentElement;
    const emptyStateClass = 'chart-empty-message';
    const topSelect = document.getElementById('corredorTopSelect');
    const orderSelect = document.getElementById('corredorOrderSelect');

    const topN = Number.isFinite(options.topN) ? options.topN : parseInt(topSelect?.value || '15', 10);
    const orderBy = options.orderBy || (orderSelect?.value || 'viajes');
    const locale = options.locale || 'es-UY';
    const currency = options.currency || 'USD';
    const currencyDigits = options.currencyDigits ?? 0;

    // Ajusta locale/currency para adaptar la moneda del dashboard.
    const currencyFormatter = new Intl.NumberFormat(locale, {
        style: 'currency',
        currency,
        minimumFractionDigits: currencyDigits,
        maximumFractionDigits: currencyDigits
    });

    const percentFormatter = new Intl.NumberFormat(locale, {
        style: 'percent',
        minimumFractionDigits: 1,
        maximumFractionDigits: 1
    });

    let dataset = prepareCorredorTarifaDataset(rows);

    if (!dataset.length) {
        if (chartCorredorTarifa) {
            chartCorredorTarifa.destroy();
            chartCorredorTarifa = null;
        }
        if (wrapper) {
            wrapper.style.height = '240px';
            canvas.style.display = 'none';
            if (!wrapper.querySelector(`.${emptyStateClass}`)) {
                const placeholder = document.createElement('div');
                placeholder.className = `loading ${emptyStateClass}`;
                placeholder.textContent = 'No hay datos de corredores para este período';
                wrapper.appendChild(placeholder);
            }
        }
        return;
    }

    if (wrapper) {
        canvas.style.display = '';
        const placeholder = wrapper.querySelector(`.${emptyStateClass}`);
        if (placeholder) {
            placeholder.remove();
        }
    }

    dataset.sort((a, b) => {
        if (orderBy === 'ventas') {
            return b.ventaTotal - a.ventaTotal;
        }
        return b.viajes - a.viajes;
    });

    const limit = Number.isFinite(topN) && topN > 0 ? topN : 15;
    dataset = dataset.slice(0, limit);

    if (wrapper) {
        const dynamicHeight = Math.max(320, dataset.length * 34);
        wrapper.style.height = `${dynamicHeight}px`;
    }

    const styles = getComputedStyle(document.documentElement);
    const positiveColor = (styles.getPropertyValue('--color-positive') || '#16a34a').trim() || '#16a34a';
    const negativeColor = (styles.getPropertyValue('--color-negative') || '#dc2626').trim() || '#dc2626';
    const targetColor = (styles.getPropertyValue('--color-accent-strong') || '#1d4ed8').trim() || '#1d4ed8';
    const textColor = (styles.getPropertyValue('--color-text-primary') || '#1e293b').trim() || '#1e293b';
    const gridColor = (styles.getPropertyValue('--color-border') || 'rgba(148, 163, 184, 0.3)').trim() || 'rgba(148, 163, 184, 0.3)';
    const surfaceColor = (styles.getPropertyValue('--color-surface') || '#ffffff').trim() || '#ffffff';

    const labels = dataset.map(item => item.corredor);
    const barValues = dataset.map(item => item.tarifaCotizadaProm);
    const targetValues = dataset.map(item => item.tarifaDeseada);
    const barColors = dataset.map(item => (item.gap >= 0 ? positiveColor : negativeColor));

    if (chartCorredorTarifa) {
        chartCorredorTarifa.destroy();
    }

    chartCorredorTarifa = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    type: 'bar',
                    label: 'Tarifa cotizada',
                    data: barValues,
                    backgroundColor: barColors,
                    borderColor: barColors,
                    borderWidth: 0,
                    borderRadius: 6,
                    borderSkipped: false,
                    order: 1
                },
                {
                    type: 'line',
                    label: 'Tarifa objetivo (20%)',
                    data: targetValues,
                    borderColor: targetColor,
                    borderWidth: 2,
                    pointBackgroundColor: targetColor,
                    pointBorderColor: targetColor,
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    tension: 0.2,
                    fill: false,
                    order: 0
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            animation: {
                duration: 650,
                easing: 'easeOutCubic'
            },
            plugins: {
                legend: {
                    labels: {
                        usePointStyle: true,
                        color: textColor
                    }
                },
                tooltip: {
                    backgroundColor: surfaceColor,
                    titleColor: textColor,
                    bodyColor: textColor,
                    borderColor: gridColor,
                    borderWidth: 1,
                    callbacks: {
                        label(context) {
                            const info = dataset[context.dataIndex];
                            if (!info) return '';
                            if (context.dataset.type === 'line') {
                                return `Tarifa deseada: ${currencyFormatter.format(info.tarifaDeseada)}`;
                            }
                            return `Tarifa cotizada: ${currencyFormatter.format(info.tarifaCotizadaProm)}`;
                        },
                        afterBody(items) {
                            if (!items.length) return [];
                            const info = dataset[items[0].dataIndex];
                            if (!info) return [];
                            const lines = [
                                `Viajes: ${info.viajes}`,
                                `Venta Flete promedio: ${currencyFormatter.format(info.ventaPromedio)}`,
                                `Tarifa deseada (80%): ${currencyFormatter.format(info.tarifaDeseada)}`,
                                `Gap: ${currencyFormatter.format(info.gap)} (${percentFormatter.format(info.gapPct || 0)})`,
                                `Total ventas: ${currencyFormatter.format(info.ventaTotal)}`,
                                `Total tarifa cotizada: ${currencyFormatter.format(info.tarifaCotizadaTotal)}`
                            ];
                            if (info.costoTipico) {
                                lines.splice(2, 0, `Costo total típico (p50): ${currencyFormatter.format(info.costoTipico)}`);
                            }
                            if (info.costDataIncomplete) {
                                lines.push('⚠ Datos de costos incompletos');
                            }
                            return lines;
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        color: gridColor
                    },
                    ticks: {
                        color: textColor,
                        callback(value) {
                            return currencyFormatter.format(value);
                        }
                    }
                },
                y: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: textColor,
                        callback(value) {
                            if (typeof value !== 'string') return value;
                            return value.length > 32 ? `${value.slice(0, 29)}…` : value;
                        }
                    }
                }
            }
        }
    });
}
  // ============================================================================
// FUNCIÓN: Renderizar tabla de operaciones
// ============================================================================
let currentSort = { column: null, direction: 'asc' };
  function renderTable() {
    const wrapper = document.getElementById('tableWrapper');
    
    if (filteredData.length === 0) {
        wrapper.innerHTML = '<div class="loading">No hay datos que coincidan con los filtros</div>';
        return;
    }
      // Ordenar datos si hay una columna seleccionada
    let sortedData = [...filteredData];
    if (currentSort.column) {
        sortedData.sort((a, b) => {
            let valA = a[currentSort.column];
            let valB = b[currentSort.column];
            
            if (typeof valA === 'string') valA = valA.toLowerCase();
            if (typeof valB === 'string') valB = valB.toLowerCase();
            
            if (valA < valB) return currentSort.direction === 'asc' ? -1 : 1;
            if (valA > valB) return currentSort.direction === 'asc' ? 1 : -1;
            return 0;
        });
    }
      const html = `
        <table id="operationsTable">
            <thead>
                <tr>
                    <th class="sortable" onclick="sortTable('fecha')">Fecha<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('carpeta')">N° Carpeta<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('contenedor')">Contenedor<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('cliente')">Cliente<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('origen')">Origen<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('destino')">Destino<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('tipoOP')">Tipo OP<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('terciarizado')">Terciarizado<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('ventaFlete')">Venta Flete<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('costosExtra')">Servicios Extra<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('totalCostos')">Costo<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('margen')">Margen<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('margenPct')">Margen %<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('horas')">Horas<div class="resizer"></div></th>
                    <th class="sortable" onclick="sortTable('kilometros')">KM<div class="resizer"></div></th>
                </tr>
            </thead>
            <tbody>
                ${sortedData.map(row => `
                    <tr>
                        <td>${formatDate(row.fecha, row.planillaMes || currentMonth)}</td>
                        <td>${row.carpeta}</td>
                        <td>${row.contenedor}</td>
                        <td>${row.cliente}</td>
                        <td>${row.origen}</td>
                        <td>${row.destino}</td>
                        <td>${row.tipoOP}</td>
                        <td>${row.terciarizado}</td>
                        <td>${formatCurrency(row.ventaFlete)}</td>
                        <td>${formatCurrency(row.costosExtra)}</td>
                        <td>${formatCurrency(row.totalCostos)}</td>
                        <td style="color: ${row.margen >= 0 ? '#10b981' : '#ef4444'}">
                            ${formatCurrency(row.margen)}
                        </td>
                        <td style="color: ${row.margenPct >= 0 ? '#10b981' : '#ef4444'}">
                            ${row.margenPct.toFixed(1)}%
                        </td>
                        <td>${row.horas}</td>
                        <td>${normalizeKey(row.terciarizado) === 'si' ? 'N/A' : row.kilometros.toFixed(0)}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
    `;
    
    wrapper.innerHTML = html;
    
    // Inicializar resize de columnas
    initColumnResize();
    
    // Actualizar indicadores de ordenamiento
    document.querySelectorAll('th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
    });
    
    if (currentSort.column) {
        const th = document.querySelector(`th[onclick="sortTable('${currentSort.column}')"]`);
        if (th) {
            th.classList.add(currentSort.direction === 'asc' ? 'sort-asc' : 'sort-desc');
        }
    }
}
  // ============================================================================
// FUNCIÓN: Inicializar resize de columnas
// ============================================================================
function initColumnResize() {
    const table = document.getElementById('operationsTable');
    if (!table) return;
    
    const ths = table.querySelectorAll('th');
    
    ths.forEach(th => {
        const resizer = th.querySelector('.resizer');
        if (!resizer) return;
        
        let startX, startWidth;
        
        resizer.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            startX = e.pageX;
            startWidth = th.offsetWidth;
            
            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        });
        
        function handleMouseMove(e) {
            const width = startWidth + (e.pageX - startX);
            if (width > 50) {
                th.style.width = width + 'px';
            }
        }
        
        function handleMouseUp() {
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        }
    });
}
  // ============================================================================
// FUNCIÓN: Ordenar tabla por columna
// ============================================================================
function sortTable(column) {
    if (currentSort.column === column) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
    } else {
        currentSort.column = column;
        currentSort.direction = 'asc';
    }
    renderTable();
}
  // ============================================================================
// FUNCIONES AUXILIARES: Formateo
// ============================================================================
function formatCurrency(value, decimals = 0) {
    const numericValue = toNumber(value);
    return '$' + numericValue.toLocaleString('en-US', {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals
    });
}
  function parseDate(value, referenceMonth) {
    if (!value) return null;
    if (value instanceof Date && !isNaN(value.getTime())) return value;
      if (typeof value === 'number') {
        const excelEpoch = new Date(Date.UTC(1899, 11, 30));
        const parsed = new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
        return isNaN(parsed.getTime()) ? null : parsed;
    }
      const str = value.toString().trim();
    if (!str) return null;
      const isoLike = str.replace(/^(\d{2})[\/.-](\d{2})[\/.-](\d{4})(.*)$/,
        (match, d, m, y, rest) => `${y}-${m}-${d}${rest}`);
    const parsed = new Date(isoLike);
    if (!isNaN(parsed.getTime())) return parsed;
      const parts = str.split(/[\/.-]/).filter(Boolean);
    if (parts.length === 3) {
        const [first, second, third] = parts.map(Number);
        if (third > 999) {
            const day = first;
            const month = second - 1;
            const year = third;
            const alt = new Date(year, month, day);
            return isNaN(alt.getTime()) ? null : alt;
        }
    }
      if (parts.length === 2) {
        const [day, month] = parts.map(Number);
        if (!isNaN(day) && !isNaN(month)) {
            let year = new Date().getFullYear();
            let refMonthNumber = null;
              if (referenceMonth) {
                const key = getMonthKeyFromSheet(referenceMonth) || referenceMonth;
                if (key && key.includes('-')) {
                    const [yearStr, monthStr] = key.split('-');
                    const parsedYear = parseInt(yearStr, 10);
                    const parsedMonth = parseInt(monthStr, 10);
                    if (!isNaN(parsedYear)) year = parsedYear;
                    if (!isNaN(parsedMonth)) refMonthNumber = parsedMonth;
                }
            }
              let monthIndex = month - 1;
            if (isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) {
                monthIndex = refMonthNumber !== null ? refMonthNumber - 1 : 0;
            }
              if (refMonthNumber !== null) {
                if (monthIndex === 11 && refMonthNumber === 1) {
                    year -= 1; // fecha de diciembre para planilla de enero
                } else if (monthIndex === 0 && refMonthNumber === 12) {
                    year += 1; // fecha de enero para planilla de diciembre
                } else if (monthIndex !== refMonthNumber - 1) {
                    // Si la fecha no coincide con la planilla, mantenemos el año de referencia
                    monthIndex = refMonthNumber - 1;
                }
            }
              const candidate = new Date(year, monthIndex, day);
            if (!isNaN(candidate.getTime())) {
                return candidate;
            }
        }
    }
      return null;
}
  function formatDate(dateStr, referenceMonth) {
    const date = parseDate(dateStr, referenceMonth);
    if (!date) return '';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${day}/${month}`;
}
  // ============================================================================
// EVENT LISTENERS: Búsquedas en tiempo real y cambio de mes
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
    // Primero, cargar dinámicamente los meses disponibles
    loadAvailableMonths();
    
    document.getElementById('searchCarpeta').addEventListener('input', applyFilters);
    document.getElementById('searchContenedor').addEventListener('input', applyFilters);
    document.getElementById('filterMes').addEventListener('change', () => {
        console.log('🔄 Cambio de mes detectado, recargando datos...');
        // Limpiar búsquedas de texto
        document.getElementById('searchCarpeta').value = '';
        document.getElementById('searchContenedor').value = '';
        // Recargar datos del nuevo mes
        fetchData();
    });
    const corredorTopSelect = document.getElementById('corredorTopSelect');
    if (corredorTopSelect) {
        corredorTopSelect.addEventListener('change', () => renderTarifaVsObjetivoPorCorredor(filteredData));
    }
    const corredorOrderSelect = document.getElementById('corredorOrderSelect');
    if (corredorOrderSelect) {
        corredorOrderSelect.addEventListener('change', () => renderTarifaVsObjetivoPorCorredor(filteredData));
    }
});
  // ============================================================================
// FUNCIÓN: Cargar dinámicamente los meses disponibles
// ============================================================================
function loadAvailableMonths() {
    console.log('🔍 Buscando planillas disponibles...');
    
    const monthNames = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
    const currentYear = new Date().getFullYear();
    const yearsToCheck = [currentYear - 1, currentYear, currentYear + 1]; // Años anteriores, actual y siguiente
    const availableSheets = [];
    let totalRequests = 0;
    let completedRequests = 0;
    
    // Generar combinaciones de mes-año
    const combinations = [];
    yearsToCheck.forEach(year => {
        monthNames.forEach(month => {
            const yearShort = year.toString().slice(-2);
            combinations.push({
                label: `${month} ${year}`,
                value: `${month} ${yearShort}`
            });
        });
    });
    
    totalRequests = combinations.length;
    
    // Intentar cargar cada combinación de forma silenciosa
    combinations.forEach(combo => {
        requestSheetData(
            combo.value,
            (data) => {
                // Verificar que data sea un array válido
                if (Array.isArray(data) && data.length > 0) {
                    console.log(`✅ Planilla disponible: ${combo.value}`);
                    availableSheets.push(combo);
                }
                completedRequests++;
                
                // Cuando terminen todas las solicitudes, actualizar el dropdown
                if (completedRequests === totalRequests) {
                    populateMonthDropdown(availableSheets);
                }
            },
            () => {
                // Silenciosamente ignorar errores de planillas no disponibles
                completedRequests++;
                if (completedRequests === totalRequests) {
                    populateMonthDropdown(availableSheets);
                }
            },
            { silent: true }
        );
    });
}

function populateMonthDropdown(availableSheets) {
    const select = document.getElementById('filterMes');
    if (!select) return;
    
    // Limpiar opciones existentes excepto "Todos"
    while (select.options.length > 1) {
        select.remove(1);
    }
    
    // Agregar opciones dinámicamente (ordenadas por fecha más reciente primero)
    availableSheets.sort((a, b) => b.value.localeCompare(a.value));
    
    availableSheets.forEach(sheet => {
        const option = document.createElement('option');
        option.value = sheet.value;
        option.textContent = sheet.label;
        select.appendChild(option);
    });
    
    console.log(`📊 Se encontraron ${availableSheets.length} planillas disponibles`);
    
    // Cargar datos con el primer mes disponible
    if (availableSheets.length > 0) {
        select.value = availableSheets[0].value;
        fetchData();
    }
}
  // ============================================================================
// INICIALIZACIÓN: Cargar datos al iniciar
// ============================================================================
// Los datos se cargarán automáticamente después de que loadAvailableMonths() complete



