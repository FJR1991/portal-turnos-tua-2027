/**
 * TUA 2027 - Datos de personal para la ficha HR del portal.
 *
 * Se pega en el Sheet "TUA 2027 - Personal" (Extensiones > Apps Script) y se
 * publica como aplicación web:
 *   Implementar > Nueva implementación > Tipo: Aplicación web
 *   Ejecutar como: Yo        Quién tiene acceso: Cualquier usuario
 *
 * El Sheet sigue siendo privado: el script lo lee con tu cuenta y solo
 * devuelve los datos si recibe la contraseña de la ficha HR. La comprobación
 * se hace aquí, en los servidores de Google, no en el navegador.
 */

// SHA-256 de la contraseña de la ficha HR. Si cambiáis la contraseña, hay que
// cambiar este valor (y el del portal).
var PASSWORD_SHA256 = 'ce512dbba6d0ce6c832884ba56d555db8f67670d667b56549037deadf0499677';

// Tras 30 contraseñas incorrectas en 10 minutos se bloquea 10 minutos.
var MAX_FALLOS = 30;

function doGet(e) {
  var p = (e && e.parameter) || {};
  var callback = String(p.callback || '').replace(/[^\w.$]/g, '');
  var result;
  try {
    result = responder(String(p.key || ''));
  } catch (err) {
    result = { ok: false, error: 'Error leyendo el listado: ' + err.message };
  }
  var body = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function responder(key) {
  var cache = CacheService.getScriptCache();
  var fallos = Number(cache.get('fallos') || 0);
  if (fallos >= MAX_FALLOS) return { ok: false, error: 'Demasiados intentos. Prueba dentro de unos minutos.' };
  if (sha256(key) !== PASSWORD_SHA256) {
    cache.put('fallos', String(fallos + 1), 600);
    return { ok: false, error: 'Contraseña incorrecta.' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return { ok: true, personal: leer(ss, 'PERSONAL'), vehiculos: leer(ss, 'VEHICULOS'), generado: new Date().toISOString() };
}

// Devuelve cada fila como objeto {CABECERA: valor}, saltando las vacías.
function leer(ss, nombre) {
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) return [];
  var valores = hoja.getDataRange().getDisplayValues();
  if (!valores.length) return [];
  var cab = valores[0].map(function (h) { return String(h).trim(); });
  return valores.slice(1).filter(function (fila) {
    return fila.some(function (v) { return String(v).trim() !== ''; });
  }).map(function (fila) {
    var o = {};
    cab.forEach(function (h, i) { if (h) o[h] = String(fila[i]).trim(); });
    return o;
  });
}

function sha256(texto) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

// Para probar desde el editor: Ejecutar > probar (mira el registro).
function probar() {
  var r = responder('PON_AQUI_LA_CONTRASEÑA');
  Logger.log(r.ok ? ('OK: ' + r.personal.length + ' personas, ' + r.vehiculos.length + ' vehículos') : r.error);
}
