/**
 * Panel de gestión: el mismo HTML se abre como barra lateral en la hoja
 * y como web (para el móvil).
 */

function abrirPanel() {
  var html = HtmlService.createTemplateFromFile('PanelUI');
  html.modo = 'lateral';
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('Panel · Ecosistema'));
}

function doGet() {
  var html = HtmlService.createTemplateFromFile('PanelUI');
  html.modo = 'web';
  return html.evaluate().setTitle('Panel · ' + cfg_().nombre)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function enlaceMovil() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert('Panel en el móvil', 'Todavía no está publicado como web. Extensiones › Apps Script › Implementar › Nueva implementación › Aplicación web (Ejecutar como: usuario que accede · Acceso: cualquier usuario con cuenta de Google).', ui.ButtonSet.OK);
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px system-ui;padding:8px">Abre este enlace en el móvil (con tu cuenta de Google) y añádelo a la pantalla de inicio. Cada persona entra con su cuenta y solo ve las hojas que tiene compartidas:<br><br>' +
    '<a href="' + url + '" target="_blank" style="word-break:break-all">' + url + '</a></div>').setWidth(460).setHeight(170);
  ui.showModalDialog(html, 'Panel en el móvil');
}

/** Todo lo que pinta la pestaña Resumen del panel para un mes. */
function datosPanelLocal_(mes) {
  var cfg = cfg_();
  var anual = resumenAnual_();
  if (!mes) {
    var hoy = new Date(), tope = cfg.anio === hoy.getFullYear() ? hoy.getMonth() : 11;
    mes = L.MESES[tope];
    for (var k = tope; k >= 0; k--) { if (anual.meses[k].conDatos) { mes = L.MESES[k]; break; } }
  }
  var idx = L.MESES.indexOf(mes);
  var actual = anual.meses[idx], previo = idx > 0 ? anual.meses[idx - 1] : null;

  var alertas = [];
  var est = actual.estado;
  if (actual.conDatos) {
    if (String(est.prime).indexOf('🔴') === 0) alertas.push({ nivel: 'alta', texto: 'Prime cost al ' + actual.kpi.prime + ' % (alerta a partir del ' + Math.round(cfg.umbrales.prime[1] * 100) + ' %).' });
    if (String(est.foodCost).indexOf('🔴') === 0) alertas.push({ nivel: 'alta', texto: 'Food cost al ' + actual.kpi.foodCost + ' %.' });
    if (String(est.personal).indexOf('🔴') === 0) alertas.push({ nivel: 'alta', texto: 'Personal al ' + actual.kpi.personal + ' % de las ventas.' });
    if (String(est.margen).indexOf('🔴') === 0) alertas.push({ nivel: 'alta', texto: 'Resultado operativo en pérdidas (' + fmtEur_(actual.resultado) + ').' });
    else if (String(est.margen).indexOf('⚠️') === 0) alertas.push({ nivel: 'media', texto: 'Margen operativo del ' + actual.kpi.margen + ' %, por debajo del objetivo.' });
    if (actual.pendiente) alertas.push({ nivel: 'media', texto: fmtEur_(actual.pendiente) + ' en facturas pendientes de pago.' });
    if (actual.diferencia && Math.abs(actual.diferencia) >= 1) alertas.push({ nivel: 'alta', texto: 'Descuadre de caja de ' + fmtEur_(actual.diferencia) + '.' });
    if (actual.gastosSinCategoria) alertas.push({ nivel: 'media', texto: fmtEur_(actual.gastosSinCategoria) + ' en gastos sin categoría válida.' });
    var g = gastos_(mes), ing = ingresos_(mes);
    var malDia = g.concat(ing).filter(function (f) { return f.dia !== '' && !(f.dia >= 1 && f.dia <= 31); }).length;
    if (malDia) alertas.push({ nivel: 'media', texto: malDia + ' líneas con día no válido.' });
    var ingSinDia = ing.filter(function (f) { return f.dia === '' && f.bruto; }).length;
    if (ingSinDia) alertas.push({ nivel: 'info', texto: ingSinDia + ' ventas sin día apuntado: no se calcula la media diaria ni la venta necesaria por día.' });
    var sinDia = g.filter(function (f) { return f.dia === ''; }).length;
    if (sinDia) alertas.push({ nivel: 'info', texto: sinDia + ' gastos sin día apuntado.' });
  }

  return {
    negocio: cfg.nombre, anio: cfg.anio, mes: mes, meses: L.MESES,
    mesesConDatos: anual.meses.filter(function (m) { return m.conDatos; }).map(function (m) { return m.mes; }),
    actual: actual, previo: previo, total: anual.total, alertas: alertas,
    serie: anual.meses.map(function (m) { return { mes: m.mes.slice(0, 3), ventas: m.ventas || 0, resultado: m.resultado || 0, prime: m.kpi.prime }; }),
    umbrales: cfg.umbrales, ia: estadoClave().configurada,
    categorias: cfg.categorias.map(function (c) { return c.categoria; }), formasPago: cfg.formasPago
  };
}

function fmtEur_(v) {
  return Utilities.formatString('%s €', (Math.round((v || 0) * 100) / 100).toFixed(2))
    .replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function include_(nombre) { return HtmlService.createHtmlOutputFromFile(nombre).getContent(); }

/** Pestaña Ventas del panel. */
function datosVentas(mes, localId) {
  var r = localId === 'todos' ? ventasTodos_(mes) : enLocal_(localId, function () {
    var v = ventasTpvMes_(mes);
    v.listaPlatos = platos_().map(function (p) { return p.plato; }).sort();
    return v;
  });
  r.localId = localId || SpreadsheetApp.getActive().getId();
  return r;
}

/** Resumen del panel para un local (id de su hoja), o 'todos'. */
function datosPanel(mes, localId) {
  var d = localId === 'todos' ? panelTodos_(mes) : enLocal_(localId, function () { return datosPanelLocal_(mes); });
  d.locales = locales_().map(function (x) { return { id: x.id, nombre: x.nombre }; });
  d.localId = localId || SpreadsheetApp.getActive().getId();
  return d;
}

/* Acciones del panel sobre el local elegido */
function importarInformeEn(archivo, localId) { return enLocal_(localId, function () { return importarInforme(archivo); }); }
function asignarNombresTpvEn(pares, localId) { return enLocal_(localId, function () { return asignarNombresTpv(pares); }); }
function guardarFacturaEn(p, localId) { return enLocal_(localId, function () { return guardarFactura(p); }); }
