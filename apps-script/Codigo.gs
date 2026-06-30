/**
 * Apps Script UNIFICADO de la planilla de Cronogramas.
 *
 * Contiene:
 *   1) La app web (doGet/doPost) que permite editar CHOFER, MATRICULA y los
 *      horarios desde las tarjetas de la app.
 *   2) Tu trigger instalable onEditInstalable (edición manual en la planilla).
 *   3) La lógica común que sincroniza la hoja "Estado" de la Planilla 2 cuando
 *      se carga una MATRICULA — funciona tanto editando a mano como desde la app.
 *
 * CÓMO PUBLICAR / ACTUALIZAR (¡importante!):
 *   1. Abrí el proyecto de Apps Script de la planilla (Extensiones → Apps Script).
 *   2. Reemplazá TODO el código por este archivo. Guardá (Ctrl+S).
 *   3. Implementar → Administrar implementaciones → (la de tipo "Aplicación web")
 *      → ícono lápiz (Editar) → en "Versión" elegí **Nueva versión** → Implementar.
 *      ⚠️ Si no creás una versión nueva, la URL /exec sigue ejecutando el código viejo.
 *      Config: Ejecutar como "Yo" · Quién accede "Cualquier usuario".
 *   4. La URL /exec es la que va en la app (⚙️ → "Guardar cambios (Apps Script)").
 *      Si la URL no cambió, no hace falta volver a pegarla.
 *
 * El trigger instalable onEditInstalable se configura una sola vez en
 * "Activadores" (reloj ⏰) → Agregar activador → función onEditInstalable,
 * evento "Al editar". (Si ya lo tenías, no hace falta tocarlo.)
 */

var PLANILLA2_ID = "121UueMo_SABwg2VE7m6Js-ctDxyfaiyt1OrS10wRZGo";
var SECRET       = ''; // dejá '' para no usar token, o poné una clave y la misma en la app

// ───────────────────────────────────────────────
// 1) APP WEB (guardar desde las tarjetas)
// ───────────────────────────────────────────────
function doGet() {
  return _json({ ok: true, msg: 'online' });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (SECRET && body.token !== SECRET) throw new Error('Token inválido');
    if (!body.spreadsheetId) throw new Error('Falta spreadsheetId');
    if (!body.tab)           throw new Error('Falta tab');

    var ss = SpreadsheetApp.openById(body.spreadsheetId);
    var sh = ss.getSheetByName(body.tab);
    if (!sh) throw new Error('No existe la pestaña "' + body.tab + '"');

    var values  = sh.getDataRange().getValues();
    var headers = values[0];
    var keyIdx  = headers.indexOf(body.keyCol);
    var colIdx  = headers.indexOf(body.col);
    if (keyIdx < 0) throw new Error('No existe la columna clave "' + body.keyCol + '"');
    if (colIdx < 0) throw new Error('No existe la columna "' + body.col + '"');

    var target = String(body.keyVal).trim();
    var rowNum = -1;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][keyIdx]).trim() === target) { rowNum = i + 1; break; }
    }
    if (rowNum < 0) throw new Error('No se encontró la fila con ' + body.keyCol + ' = ' + body.keyVal);

    sh.getRange(rowNum, colIdx + 1).setValue(body.value);

    // Si se editó la MATRICULA, replicamos la sincronización con la Planilla 2.
    actualizarEstadoDesdeMatricula(body.tab, sh, rowNum, colIdx + 1);

    return _json({ ok: true, row: rowNum, col: body.col, value: body.value });
  } catch (err) {
    return _json({ ok: false, error: String(err && err.message || err) });
  }
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────────
// 2) TRIGGER INSTALABLE (edición manual en la planilla)
// ───────────────────────────────────────────────
function onEditInstalable(e) {
  var range = e.range;
  var sheet = range.getSheet();
  actualizarEstadoDesdeMatricula(sheet.getName(), sheet, range.getRow(), range.getColumn());
}

// ───────────────────────────────────────────────
// 3) LÓGICA COMÚN: al cargar una MATRICULA, actualiza "Estado" en Planilla 2.
//    Se usa tanto desde el trigger manual como desde la app web (doPost).
//    'col' es la columna (1-based) que se acaba de editar.
// ───────────────────────────────────────────────
function actualizarEstadoDesdeMatricula(sheetName, sheet, row, col) {
  var matriculaCol, contenedorCol, destinoCol;

  if (sheetName === "FORD" && col === 9) {
    matriculaCol  = 9;
    contenedorCol = 2;
    destinoCol    = 7;
  } else if (sheetName === "PEUGEOT" && col === 10) {
    matriculaCol  = 10;
    contenedorCol = 4;
    destinoCol    = 8;
  } else {
    return;
  }

  if (row === 1) return;

  var matricula = sheet.getRange(row, matriculaCol).getValue();
  if (!matricula || String(matricula).trim() === "") return;

  var contenedor = sheet.getRange(row, contenedorCol).getValue();
  var destino    = sheet.getRange(row, destinoCol).getValue();
  var cliente    = sheetName + " / " + destino;

  try {
    var planilla2   = SpreadsheetApp.openById(PLANILLA2_ID);
    var estadoSheet = planilla2.getSheetByName("Estado");

    if (!estadoSheet) {
      Logger.log("ERROR: No se encontro la hoja Estado en Planilla 2.");
      return;
    }

    var data         = estadoSheet.getDataRange().getValues();
    var matriculaStr = String(matricula).trim().toUpperCase();
    var encontrado   = false;

    for (var i = 1; i < data.length; i++) {
      var valorC = String(data[i][2]).trim().toUpperCase();
      if (valorC === matriculaStr) {
        var targetRow = i + 1;
        var now  = new Date();
        var hora = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
        estadoSheet.getRange(targetRow, 4).setValue("Asignado");
        estadoSheet.getRange(targetRow, 5).setValue(hora);
        estadoSheet.getRange(targetRow, 6).setValue(cliente);
        estadoSheet.getRange(targetRow, 7).setValue(contenedor);
        Logger.log("Actualizado fila " + targetRow + " mat:" + matricula + " cliente:" + cliente + " cont:" + contenedor);
        encontrado = true;
        break;
      }
    }

    if (!encontrado) {
      Logger.log("AVISO: Matricula " + matricula + " no encontrada en columna C.");
    }

  } catch (err) {
    Logger.log("ERROR: " + err.message);
  }
}
