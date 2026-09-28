/**
 * Conexión con la API de Claude (HTTP directo: Apps Script no tiene SDK oficial).
 * La clave se guarda en las propiedades privadas del usuario, nunca en la hoja.
 */
var CLAUDE_URL = 'https://api.anthropic.com/v1/messages';
var MODELO_POR_DEFECTO = 'claude-opus-5';

/** Clave del usuario; si no tiene, la compartida del local (la que el dueño decidió compartir con su equipo). */
function claveApi_() {
  var k = PropertiesService.getUserProperties().getProperty('ANTHROPIC_API_KEY') ||
          PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY_LOCAL');
  if (!k) throw new Error('Falta la clave de la IA. Ve al menú 🍽️ Ecosistema › 🔑 Configurar clave de IA.');
  return k;
}
function modeloIA_() {
  return PropertiesService.getDocumentProperties().getProperty('MODELO_IA') || MODELO_POR_DEFECTO;
}

/** Llama a /v1/messages. Reintenta una vez ante saturación o errores del servidor. */
function claude_(body) {
  body.model = body.model || modeloIA_();
  var headers = { 'x-api-key': claveApi_(), 'anthropic-version': '2023-06-01' };
  if (body.model === 'claude-opus-5') {
    // Respaldo en el servidor: si el modelo rechaza la petición, la API la repite con el modelo recomendado.
    body.fallbacks = 'default';
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
  }
  var opts = { method: 'post', contentType: 'application/json', headers: headers,
               payload: JSON.stringify(body), muteHttpExceptions: true };
  var res, code;
  for (var intento = 0; intento < 3; intento++) {
    res = UrlFetchApp.fetch(CLAUDE_URL, opts);
    code = res.getResponseCode();
    if (code === 429 || code === 529 || code >= 500) { Utilities.sleep(2000 * (intento + 1)); continue; }
    break;
  }
  var json;
  try { json = JSON.parse(res.getContentText()); } catch (e) { throw new Error('Respuesta no válida de la IA (' + code + ').'); }
  if (code !== 200) {
    var msg = json && json.error && json.error.message ? json.error.message : ('código ' + code);
    if (code === 401) msg = 'La clave de la IA no es válida. Vuelve a configurarla.';
    if (code === 400 && /credit balance/i.test(msg)) msg = 'Tu cuenta de Anthropic no tiene saldo. Recárgala en console.anthropic.com.';
    throw new Error('Error de la IA: ' + msg);
  }
  if (json.stop_reason === 'refusal') throw new Error('La IA no ha podido procesar esta petición. Prueba a reformularla.');
  return json;
}

function textoDe_(resp) {
  return (resp.content || []).filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('\n').trim();
}

/* ---------------------------------------------------------- clave de API */

function configurarClave() {
  var html = HtmlService.createHtmlOutputFromFile('Clave').setWidth(470).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configurar la IA');
}

/** Valida la clave con una petición mínima y la guarda. Llamada desde Clave.html. */
function guardarClave(clave, modelo, compartir) {
  clave = String(clave || '').trim();
  if (!/^sk-ant-/.test(clave)) throw new Error('La clave debe empezar por "sk-ant-".');
  var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/models?limit=1', {
    headers: { 'x-api-key': clave, 'anthropic-version': '2023-06-01' }, muteHttpExceptions: true });
  if (res.getResponseCode() === 401) throw new Error('Anthropic rechaza esa clave. Revísala.');
  if (res.getResponseCode() !== 200) throw new Error('No se pudo comprobar la clave (código ' + res.getResponseCode() + ').');
  PropertiesService.getUserProperties().setProperty('ANTHROPIC_API_KEY', clave);
  var sp = PropertiesService.getScriptProperties();
  // Solo quien compartió esa misma clave puede dejar de compartirla (evita que otra persona la borre al guardar la suya)
  if (compartir) sp.setProperty('ANTHROPIC_API_KEY_LOCAL', clave);
  else if (sp.getProperty('ANTHROPIC_API_KEY_LOCAL') === clave) sp.deleteProperty('ANTHROPIC_API_KEY_LOCAL');
  if (modelo) PropertiesService.getDocumentProperties().setProperty('MODELO_IA', modelo);
  return 'Clave guardada. Ya puedes usar el asistente.';
}

function estadoClave() {
  var k = PropertiesService.getUserProperties().getProperty('ANTHROPIC_API_KEY');
  var comp = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY_LOCAL');
  return { configurada: !!(k || comp), propia: !!k, compartida: !!comp, final: (k || comp || '').slice(-4), modelo: modeloIA_() };
}
