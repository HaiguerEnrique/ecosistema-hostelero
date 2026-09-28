/**
 * Varios locales en un solo panel.
 * Cada local tiene su hoja; la tabla "Locales del grupo" (Configuración R:S) enlaza las demás.
 * Las funciones de datos leen de libro_(): la hoja activa o la del local elegido en el panel.
 */

var LIBRO_ = null;          // hoja del local en uso durante una llamada
var LIBROS_CACHE_ = {};

function libro_() { return LIBRO_ || SpreadsheetApp.getActive(); }

/** Ejecuta fn con la hoja del local indicado (id de la hoja). Sin id → la hoja activa. */
function enLocal_(id, fn) {
  var activo = SpreadsheetApp.getActive();
  if (!id || id === activo.getId()) return fn();
  var previo = LIBRO_;
  LIBRO_ = LIBROS_CACHE_[id] || (LIBROS_CACHE_[id] = SpreadsheetApp.openById(id));
  try { return fn(); } finally { LIBRO_ = previo; }
}

/** Locales visibles desde esta hoja: ella misma primero y luego los de la tabla. */
function locales_() {
  var ss = SpreadsheetApp.getActive();
  var cfg = ss.getSheetByName(L.CFG);
  var propio = { id: ss.getId(), nombre: String(cfg.getRange(L.C.nombre).getValue() || ss.getName()) };
  var out = [propio], vistos = {}; vistos[propio.id] = true;
  cfg.getRange('R' + L.C.listFirst + ':S' + L.C.locales.last).getValues().forEach(function (f) {
    var m = String(f[1] || '').match(/\/d\/([a-zA-Z0-9_-]{20,})/) || String(f[1] || '').match(/^([a-zA-Z0-9_-]{30,})$/);
    if (!m || vistos[m[1]]) return;
    vistos[m[1]] = true;
    if (!accesible_(m[1])) return;   // quien no tenga acceso a ese local no lo ve
    out.push({ id: m[1], nombre: String(f[0] || 'Local ' + out.length) });
  });
  return out;
}

function nombreLocal_(id) {
  var l = locales_().filter(function (x) { return x.id === id; })[0];
  return l ? l.nombre : '';
}

/** Tabla de locales en Configuración (constructor y parche 2.2.0). */
function prepararLocales_(ss) {
  var cfg = ss.getSheetByName(L.CFG), f = L.C.listFirst;
  seccion_(cfg, 19, 18, 2, 'LOCALES DEL GRUPO (para el panel)');
  cabecera_(cfg, 20, 18, ['Local', 'Enlace de su hoja de control']);
  bordes_(cfg.getRange('R20:S' + L.C.locales.last));
  cfg.getRange('R' + (L.C.locales.last + 1)).setValue('Solo si tienes más de un local: pega aquí el enlace de la hoja de cada uno de los otros. El panel te dejará elegir local o verlos todos juntos.')
    .setFontColor(COLOR.note).setWrap(true);
  cfg.getRange('R' + (L.C.locales.last + 1) + ':S' + (L.C.locales.last + 1)).merge();
  cfg.setRowHeight(L.C.locales.last + 1, 45);
  cfg.setColumnWidth(17, 20); cfg.setColumnWidth(18, 160); cfg.setColumnWidth(19, 360);
}

/* --------------------------------------------------- vistas consolidadas */

function sumar_(a, b) {
  if (typeof a === 'number' || typeof b === 'number') return (typeof a === 'number' ? a : 0) + (typeof b === 'number' ? b : 0);
  if (a && b && typeof a === 'object' && typeof b === 'object' && !Array.isArray(a)) {
    var o = {}; Object.keys(a).concat(Object.keys(b)).forEach(function (k) { o[k] = sumar_(a[k], b[k]); }); return o;
  }
  return a != null && a !== '' ? a : b;
}

function recalcularKpi_(m, u) {
  var v = m.ventas || 0;
  var p = function (x) { return v ? Math.round(x / v * 1000) / 10 : null; };
  m.kpi = { foodCost: p(m.cmv || 0), personal: p(m.personal || 0), prime: p((m.cmv || 0) + (m.personal || 0)),
            margen: p(m.resultado || 0), explotacion: p(m.explotacion || 0) };
  var est = function (val, par, alto) {
    if (val == null) return '—';
    return val <= par[0] * 100 ? '✅ Bien' : val <= par[1] * 100 ? '⚠️ Vigilar' : '🔴 ' + alto;
  };
  m.estado = { foodCost: est(m.kpi.foodCost, u.foodCost, 'Alto'), personal: est(m.kpi.personal, u.personal, 'Alto'),
               prime: est(m.kpi.prime, u.prime, 'Alto'),
               margen: m.kpi.margen == null ? '—' : m.kpi.margen >= u.margen[0] * 100 ? '✅ Bien' : m.kpi.margen >= u.margen[1] * 100 ? '⚠️ Vigilar' : '🔴 Pérdidas' };
  m.ventaDia = null; m.ventaNecDia = null;   // no tiene sentido sumado entre locales
  m.conDatos = !!v || (m.gastosPorCategoria || []).length > 0;
  return m;
}

function gastosCat_(listas) {
  var acc = {};
  listas.forEach(function (l) { (l || []).forEach(function (g) {
    var x = acc[g.categoria] || (acc[g.categoria] = { categoria: g.categoria, tipo: g.tipo, importe: 0 });
    x.importe = Math.round((x.importe + g.importe) * 100) / 100;
  }); });
  return Object.keys(acc).map(function (k) { return acc[k]; });
}

/** Resumen de todos los locales sumados + comparativa. */
function panelTodos_(mes) {
  var locs = locales_();
  var partes = locs.map(function (l) { return { l: l, d: enLocal_(l.id, function () { return datosPanelLocal_(mes); }) }; });
  var base = partes[0].d, u = base.umbrales;
  var junta = function (k) {
    var r = partes.reduce(function (acc, p) { return acc ? sumar_(acc, p.d[k]) : JSON.parse(JSON.stringify(p.d[k])); }, null);
    return r;
  };
  var actual = recalcularKpi_(junta('actual'), u);
  actual.gastosPorCategoria = gastosCat_(partes.map(function (p) { return p.d.actual.gastosPorCategoria; }));
  var previo = base.previo ? recalcularKpi_(junta('previo'), u) : null;
  var total = junta('total');
  if (total.ventas) {
    total.primePct = Math.round((total.cmv + total.personal) / total.ventas * 1000) / 10;
    total.margenPct = Math.round(total.resultado / total.ventas * 1000) / 10;
  }
  var serie = base.serie.map(function (s, i) {
    return { mes: s.mes, ventas: partes.reduce(function (a, p) { return a + (p.d.serie[i].ventas || 0); }, 0),
             resultado: partes.reduce(function (a, p) { return a + (p.d.serie[i].resultado || 0); }, 0) };
  });
  var alertas = [];
  partes.forEach(function (p) { p.d.alertas.forEach(function (a) { alertas.push({ nivel: a.nivel, texto: p.l.nombre + ': ' + a.texto }); }); });
  return {
    negocio: 'Todos los locales', anio: base.anio, mes: base.mes, meses: base.meses, todos: true,
    actual: actual, previo: previo, total: total, serie: serie, alertas: alertas, umbrales: u, ia: base.ia,
    categorias: base.categorias, formasPago: base.formasPago,
    comparativa: partes.map(function (p) {
      var a = p.d.actual;
      return { local: p.l.nombre, ventas: a.ventas, resultado: a.resultado, margen: a.kpi.margen, prime: a.kpi.prime,
               foodCost: a.kpi.foodCost, personal: a.kpi.personal, caja: a.cajaGenerada,
               estado: { prime: a.estado.prime, margen: a.estado.margen, foodCost: a.estado.foodCost, personal: a.estado.personal } };
    })
  };
}

/** Ventas de todos los locales: comparativa por local. */
function ventasTodos_(mes) {
  return {
    mes: mes, todos: true, conDatos: true,
    comparativa: locales_().map(function (l) {
      var v = enLocal_(l.id, function () { return ventasTpvMes_(mes); });
      var llevar = v.canales ? v.canales['Para llevar'] + v.canales.Recoger + v.canales.Domicilio : 0;
      return { local: l.nombre, conDatos: v.conDatos, ventas: v.ventas, noches: v.noches,
               mediaNoche: v.noches ? Math.round(v.ventas / v.noches) : null, ticketMedio: v.ticketMedio, pedidos: v.pedidos,
               pctLlevar: v.ventas ? Math.round(llevar / v.ventas * 1000) / 10 : null,
               pctCancelado: v.ventas ? Math.round(v.cancelado / (v.ventas + v.cancelado) * 1000) / 10 : null,
               cmv: v.cmv };
    })
  };
}

/** ¿Puede el usuario actual abrir esa hoja? (cacheado 10 min por usuario) */
function accesible_(id) {
  var cache = CacheService.getUserCache(), k = 'acc_' + id, v = cache.get(k);
  if (v !== null) return v === '1';
  var ok = true;
  try { DriveApp.getFileById(id).getName(); } catch (e) { ok = false; }
  cache.put(k, ok ? '1' : '0', 600);
  return ok;
}
