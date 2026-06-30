/**
 * App web para guardar cambios desde la app de Cronogramas en la planilla.
 *
 * CÓMO PUBLICARLO (una sola vez):
 *  1. Abrí https://script.google.com  →  "Nuevo proyecto".
 *  2. Borrá lo que haya y pegá TODO este archivo.
 *  3. (Opcional) Poné un token en SECRET y el mismo en la app (campo "Token").
 *  4. Implementar  →  Nueva implementación  →  tipo "Aplicación web".
 *       - Ejecutar como: "Yo".
 *       - Quién tiene acceso: "Cualquier usuario".
 *  5. Copiá la URL que termina en /exec y pegala en la app
 *     (⚙️ → "Guardar cambios (Apps Script)").
 *
 * El usuario que publica debe tener permiso de EDICIÓN sobre la planilla.
 */

var SECRET = ''; // dejá '' para no usar token, o poné una clave y la misma en la app

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
