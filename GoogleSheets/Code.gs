/**
 * Apps Script para el formulario de consentimiento de La Rufina.
 *
 * - doPost: recibe el envío del formulario público (index.html), guarda
 *   los datos en la pestaña "Registros" y la firma como imagen en Drive.
 * - doGet: usado por la página interna (registros.html) para listar los
 *   registros y para descargar la imagen de una firma puntual. Requiere
 *   una clave de acceso (CLAVE_ACCESO) para no dejar los datos abiertos
 *   a cualquiera que tenga la URL.
 *
 * INSTALACIÓN / REDESPLIEGUE:
 * 1. Cambia CLAVE_ACCESO por una clave propia (ábrela solo tú y las
 *    personas de la empresa que deban ver los registros).
 * 2. Extensiones > Apps Script > pega este archivo completo en Code.gs.
 * 3. Implementar > Gestionar implementaciones > Editar (lápiz) >
 *    Nueva versión > Implementar. (Si es la primera vez: Nueva
 *    implementación, tipo "Aplicación web", ejecutar como "Yo", acceso
 *    "Cualquier usuario").
 * 4. La primera vez que uses DriveApp, Google te pedirá autorizar
 *    permisos nuevos (acceso a Drive) — acéptalos, es tu propio script.
 * 5. La URL de la aplicación web no cambia entre versiones, así que no
 *    hace falta tocar SCRIPT_URL en index.html si ya la tenías puesta.
 */

var NOMBRE_HOJA = 'Registros';
var CARPETA_FIRMAS = 'Firmas Consentimientos La Rufina';
var CLAVE_ACCESO = 'CAMBIA-ESTA-CLAVE';

// Orden fijo de columnas: se usa tanto para escribir como para leer.
var CAMPOS = [
  { header: 'ID', key: 'id' },
  { header: 'Fecha y hora de registro', key: 'timestamp' },
  { header: 'Día', key: 'dia' },
  { header: 'Mes', key: 'mes' },
  { header: 'Año', key: 'anio' },
  { header: 'Calidad', key: 'calidad' },
  { header: 'Tipo de documento', key: 'tipoDoc' },
  { header: 'Número de documento', key: 'cedula' },
  { header: 'Nombre del acudiente', key: 'nombreAcudiente' },
  { header: 'Teléfono', key: 'telefono' },
  { header: 'Nombre del menor', key: 'nombreMenor' },
  { header: 'Edad del menor', key: 'edadMenor' },
  { header: 'Fecha de visita', key: 'fechaVisita' },
  { header: 'ID Firma (Drive)', key: 'firmaId' }
];
var COL_ID = 0;
var COL_FIRMA_ID = 13;

function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    var hoja = obtenerHoja();

    var firmaId = '';
    if (datos.firmaImg) {
      try {
        firmaId = guardarFirma(datos.firmaImg, datos.id);
      } catch (errFirma) {
        Logger.log('Error al guardar firma en Drive: ' + errFirma.message);
        firmaId = '';
      }
    }

    var fila = CAMPOS.map(function(c) {
      if (c.key === 'firmaId') return firmaId;
      return datos[c.key] || '';
    });
    hoja.appendRow(fila);

    return respuestaJSON({ ok: true });
  } catch (err) {
    return respuestaJSON({ ok: false, error: err.message });
  }
}

function doGet(e) {
  var clave = e.parameter.clave || '';
  if (clave !== CLAVE_ACCESO) {
    return respuestaJSON({ ok: false, error: 'Clave incorrecta.' });
  }
  if (e.parameter.accion === 'firma') {
    return obtenerFirma(e.parameter.id);
  }
  return listarRegistros();
}

function listarRegistros() {
  var hoja = obtenerHoja();
  var datos = hoja.getDataRange().getValues();
  var registros = [];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var obj = {};
    CAMPOS.forEach(function(c, idx) {
      if (c.key === 'firmaId') {
        obj.tieneFirma = !!fila[idx];
      } else {
        obj[c.key] = formatearValor(fila[idx]);
      }
    });
    registros.push(obj);
  }
  return respuestaJSON({ ok: true, registros: registros });
}

function formatearValor(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(v);
}

function obtenerFirma(id) {
  var hoja = obtenerHoja();
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL_ID]) === String(id)) {
      var firmaId = datos[i][COL_FIRMA_ID];
      if (!firmaId) return respuestaJSON({ ok: false, error: 'Este registro no tiene firma guardada.' });
      var blob = DriveApp.getFileById(firmaId).getBlob();
      var b64 = Utilities.base64Encode(blob.getBytes());
      return respuestaJSON({ ok: true, firma: 'data:image/png;base64,' + b64 });
    }
  }
  return respuestaJSON({ ok: false, error: 'Registro no encontrado.' });
}

function guardarFirma(dataUrl, id) {
  var base64 = dataUrl.split(',')[1];
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, 'image/png', 'firma_' + id + '.png');
  var carpeta = obtenerCarpetaFirmas();
  var archivo = carpeta.createFile(blob);
  return archivo.getId();
}

function obtenerCarpetaFirmas() {
  var carpetas = DriveApp.getFoldersByName(CARPETA_FIRMAS);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(CARPETA_FIRMAS);
}

function obtenerHoja() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = libro.insertSheet(NOMBRE_HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(CAMPOS.map(function(c) { return c.header; }));
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function respuestaJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
