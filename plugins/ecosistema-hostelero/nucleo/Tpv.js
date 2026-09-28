/**
 * Informes del TPV (Foodyservice): importación y agregados.
 * - Informe de ventas (sales-report): una fila por línea vendida.
 * - Informe de facturas (invoices-report): una fila por ticket (cobros por forma de pago).
 * Todo se guarda POR DÍA (noche de servicio), así se puede subir el informe de un solo día cada noche
 * o el de todo el año: solo se sustituyen los días que trae el informe.
 * Reglas (validadas con un local real): lo vendido antes de las 6:00 cuenta para la noche anterior;
 * un pedido = misma mesa con menos de 60 min entre líneas; domicilio/recoger = cada hora de creación.
 * Los datos de clientes del informe de facturas NO se guardan.
 */

var TPV = {
  DIAS: 'TPV Días', HORAS: 'TPV Horas', CAM: 'TPV Camareros', PLATOS: 'TPV Platos (día)',
  DIAS_COLS: ['Fecha (noche)', 'Día', 'Ventas sin IVA', 'Pedidos', 'Sala', 'Terraza', 'Para llevar', 'Recoger', 'Domicilio',
              'Pedidos para llevar', 'Cancelado (sin IVA)', 'Tickets', 'Ticket medio (IVA incl.)', 'Cobrado efectivo', 'Cobrado tarjeta', 'Cobrado otros'],
  HORAS_COLS: ['Fecha (noche)', 'Día', 'Hora', 'Ventas sin IVA', 'Pedidos'],
  CAM_COLS: ['Fecha (noche)', 'Camarero', 'Ventas sin IVA', 'Pedidos', 'Cancelado (sin IVA)', 'Uds. canceladas'],
  PLATOS_COLS: ['Fecha (noche)', 'Producto en el TPV', 'Unidades', 'Uds. para llevar', 'Importe (IVA incl.)'],
  DOW: ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'],
  MARCA: 'Auto · TPV',          // en Observaciones de las líneas de ingresos creadas desde el TPV
  ORIGEN: 'TPV'                 // columna Origen de Ventas TPV para lo reconstruido desde los días
};

/* ------------------------------------------------------------- importar */

/** Llamada desde el panel. archivo = {nombre, mime, base64}. */
function importarInforme(archivo) {
  var bytes = Utilities.base64Decode(archivo.base64);
  var blob = Utilities.newBlob(bytes, archivo.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', archivo.nombre);
  var filas = leerTabla_(blob);
  if (!filas.length) throw new Error('El archivo está vacío.');
  var cab = filas[0].map(function (h) { return String(h).trim(); });
  var obj = filas.slice(1).map(function (f) { var o = {}; cab.forEach(function (h, i) { o[h] = f[i]; }); return o; });

  var res, agr, desde = tpvDesde_();
  var recortar = function (a) {
    recortarDesde_(a, desde);
    if (!a.lista.length) throw new Error('Todo «' + archivo.nombre + '» es anterior al ' + desde.split('-').reverse().join('/') +
      ' (Configuración › Datos del TPV desde). No se ha guardado nada.');
    return a;
  };
  if (cab.indexOf('Mesa/Pedido') >= 0 && (cab.indexOf('Fecha creacion') >= 0 || cab.indexOf('Fecha creación') >= 0)) res = guardarVentas_(agr = recortar(agregarVentas_(obj)));
  else if (cab.indexOf('Nº Factura') >= 0 && cab.indexOf('Base Imponible') >= 0) res = guardarFacturas_(agr = recortar(agregarFacturas_(obj)));
  else if (cab.indexOf('Base Imponible Total') >= 0 || cab.indexOf('Total (Con propinas)') >= 0) {
    throw new Error('«' + archivo.nombre + '» es el «Resumen de facturación» (solo totales). Hace falta el «Informe de facturas», el que trae una línea por ticket (archivo invoices-report…).');
  }
  else throw new Error('No reconozco «' + archivo.nombre + '». Sube el «Informe de ventas» o el «Informe de facturas» de Foodyservice (Excel).');
  res.archivo = archivo.nombre;
  if (agr.ignoradas) {
    res.aviso = (res.aviso ? res.aviso + ' ' : '') + 'Se han ignorado ' + agr.ignoradas + ' noches anteriores al ' +
      desde.split('-').reverse().join('/') + ' (Configuración › Datos del TPV desde).';
  }

  // Guarda el original en Informes TPV/<año>
  try {
    var c = carpetas_(true);
    c.tpv.createFile(blob).setName(Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd HH.mm') + ' · ' + archivo.nombre);
  } catch (e) { res.aviso = (res.aviso ? res.aviso + ' ' : '') + 'No se pudo guardar el original en Drive: ' + e.message; }
  return res;
}

/** Convierte el Excel a una hoja temporal para leerlo con fiabilidad y la tira a la papelera. */
function leerTabla_(blob) {
  var tmp = Drive.Files.create({ name: '_import_tpv_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, blob);
  try {
    return SpreadsheetApp.openById(tmp.id).getSheets()[0].getDataRange().getValues();
  } finally {
    DriveApp.getFileById(tmp.id).setTrashed(true);
  }
}

function num_(x) {
  if (typeof x === 'number') return x;
  var n = parseFloat(String(x == null ? '' : x).replace('%', '').replace(/\s/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}
function iso_(d) { return Utilities.formatDate(d, 'Europe/Madrid', 'yyyy-MM-dd'); }
function fechaHora_(v) {
  if (v instanceof Date) return v;
  var m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? new Date(+m[3], m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0)) : null;
}
function fecha_(v) {
  if (v instanceof Date) return iso_(v);
  var m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : null;
}
function canal_(m) {
  m = String(m || '');
  if (/^terraza/i.test(m)) return 'Terraza';
  if (/^llevar/i.test(m)) return 'Llevar';
  if (/recoger/i.test(m)) return 'Recoger';
  if (/domicilio/i.test(m)) return 'Domicilio';
  return 'Sala';
}
function esLlevar_(c) { return c === 'Llevar' || c === 'Recoger' || c === 'Domicilio'; }
function diaDe_(v) { return v instanceof Date ? iso_(v) : (String(v).match(/^\d{4}-\d{2}-\d{2}/) || [''])[0]; }
function mesDeCelda_(v) { return v instanceof Date ? iso_(v).slice(0, 7) : (String(v).match(/^\d{4}-\d{2}/) || [''])[0]; }
function aFecha_(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +(p[2] || 1)); }
function dow_(dia) { return TPV.DOW[(aFecha_(dia).getDay() + 6) % 7]; }

/** Agrega el informe de ventas por día: platos, canales, horas y camareros. */
function agregarVentas_(rows) {
  var lin = [];
  rows.forEach(function (r) {
    var f = fechaHora_(r['Fecha creacion'] || r['Fecha creación']);
    if (!f) return;
    var d = new Date(f.getTime()); if (d.getHours() < 6) d.setDate(d.getDate() - 1);
    var iva = num_(r['% IVA'] || 10); if (iva > 1) iva = iva / 100;
    var tot = num_(r['Total']);
    lin.push({ t: f.getTime(), dia: iso_(d), h: f.getHours(),
      mesa: String(r['Mesa/Pedido'] || '').trim(), c: canal_(r['Mesa/Pedido']),
      cam: String(r['Camarero'] || '').trim() || '—', extra: String(r['Tipo'] || '') === 'EXTRA',
      nombre: String(r['Nombre'] || '').replace(/\s+/g, ' ').trim(),
      u: num_(r['Cantidad']), tot: tot, e: tot / (1 + iva),
      canc: /^s/i.test(String(r['Cancelado'] || '')), reemb: /^s/i.test(String(r['Reembolsado'] || '')) });
  });
  lin.sort(function (a, b) { return a.t - b.t; });
  var ult = {}, n = 0;
  lin.forEach(function (l) {
    if (l.canc || l.reemb) return;
    var k = /[0-9]/.test(l.mesa) ? l.mesa : l.mesa + '@' + l.t;
    var u = ult[k];
    if (!u || l.t - u.t > 3600e3 || u.dia !== l.dia) ult[k] = { id: ++n, t: l.t, dia: l.dia }; else u.t = l.t;
    l.ped = ult[k].id;
  });
  var prod = {}, dias = {}, horas = {}, cams = {}, visto = {};
  var D = function (dia) {
    return dias[dia] || (dias[dia] = { e: 0, p: 0, can: { Sala: 0, Terraza: 0, Llevar: 0, Recoger: 0, Domicilio: 0 }, pl: 0, canc: 0 });
  };
  lin.forEach(function (l) {
    var ck = l.dia + '|' + l.cam, cam = cams[ck] || (cams[ck] = { dia: l.dia, cam: l.cam, e: 0, p: 0, ce: 0, cu: 0 });
    var d = D(l.dia);
    if (l.canc) { cam.ce += l.e; cam.cu += l.u; d.canc += l.e; return; }
    if (l.reemb) return;
    var nombre = (l.extra ? '+ ' : '') + l.nombre, pk = l.dia + '|' + nombre;
    var p = prod[pk] || (prod[pk] = { dia: l.dia, nombre: nombre, u: 0, ul: 0, tot: 0 });
    p.u += l.u; p.tot += l.tot; if (esLlevar_(l.c)) p.ul += l.u;
    d.e += l.e; cam.e += l.e;
    var pi = visto[l.ped];
    if (!pi) {
      pi = visto[l.ped] = { h: l.h, c: l.c };
      d.p++; cam.p++; if (esLlevar_(pi.c)) d.pl++;
      var hk = l.dia + '|' + pi.h; (horas[hk] = horas[hk] || { e: 0, p: 0 }).p++;
    }
    d.can[pi.c] += l.e;
    var hk2 = l.dia + '|' + pi.h; (horas[hk2] = horas[hk2] || { e: 0, p: 0 }).e += l.e;
  });
  var dk = Object.keys(dias).sort();
  return { dias: dias, lista: dk, meses: unicos_(dk.map(function (x) { return x.slice(0, 7); })), prod: prod, horas: horas, cams: cams };
}

function agregarFacturas_(rows) {
  var anuladas = {};
  rows.forEach(function (r) { if (/rectific/i.test(r['Tipo'] || '')) anuladas[String(r['Rectifica a'] || '').trim()] = true; });
  var dias = {};
  rows.forEach(function (r) {
    if (/rectific/i.test(r['Tipo'] || '') || anuladas[String(r['Nº Factura'] || '').trim()]) return;
    var dia = fecha_(r['Fecha']); if (!dia) return;
    var tot = num_(r['Total sin propina'] !== '' && r['Total sin propina'] != null ? r['Total sin propina'] : r['Total']);
    var met = /efectivo/i.test(r['Método de Pago'] || '') ? 'ef' : /tarjeta/i.test(r['Método de Pago'] || '') ? 'ta' : 'ot';
    var d = dias[dia] || (dias[dia] = { n: 0, tot: 0, ef: 0, ta: 0, ot: 0 });
    d.n++; d.tot += tot; d[met] += tot;
  });
  var dk = Object.keys(dias).sort();
  return { dias: dias, lista: dk, meses: unicos_(dk.map(function (x) { return x.slice(0, 7); })) };
}
function unicos_(a) { var o = {}; a.forEach(function (x) { o[x] = true; }); return Object.keys(o).sort(); }

/** Configuración › «Apuntar ingresos desde el TPV»: con «No» los informes se guardan pero los ingresos los apunta el dueño. */
function ingresosAutomaticos_() {
  return String(libro_().getSheetByName(L.CFG).getRange(L.C.ingAuto).getValue()).trim().toLowerCase() !== 'no';
}

/** Configuración › «Datos del TPV desde»: las noches anteriores se ignoran al subir informes ('' = sin límite). */
function tpvDesde_() {
  var v = libro_().getSheetByName(L.CFG).getRange(L.C.tpvDesde).getValue();
  return v instanceof Date ? iso_(v) : '';
}
/** Quita de un informe ya agregado las noches anteriores a `desde`. Deja en a.ignoradas cuántas quitó. */
function recortarDesde_(a, desde) {
  a.ignoradas = 0;
  if (!desde) return a;
  var fuera = a.lista.filter(function (d) { return d < desde; });
  if (!fuera.length) return a;
  fuera.forEach(function (d) { delete a.dias[d]; });
  a.lista = a.lista.filter(function (d) { return d >= desde; });
  a.meses = unicos_(a.lista.map(function (x) { return x.slice(0, 7); }));
  ['prod', 'horas', 'cams'].forEach(function (k) {
    if (a[k]) Object.keys(a[k]).forEach(function (key) { if (key.slice(0, 10) < desde) delete a[k][key]; });
  });
  a.ignoradas = fuera.length;
  return a;
}

/* -------------------------------------------------------------- guardar */

function hojaDatos_(nombre, cols) {
  var ss = libro_(), sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre, ss.getSheets().length - 1);   // antes de Configuración
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#d5e4de').setWrap(true);
    sh.setFrozenRows(1); sh.setTabColor('#8a8f8d');
    sh.getRange(1, 1, 1, cols.length).protect().setDescription('Cabecera de datos del TPV').setWarningOnly(true);
    sh.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
  }
  return sh;
}

function leerDatos_(sh, ncol) {
  return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, ncol).getValues().filter(function (f) { return f[0] !== ''; }) : [];
}

/** Sustituye en una hoja de datos las filas de los días `dias` por `nuevas` (col A = fecha). */
function reemplazarDias_(sh, ncol, dias, nuevas, orden) {
  var set = {}; dias.forEach(function (d) { set[d] = true; });
  var viejas = leerDatos_(sh, ncol);
  var todas = viejas.filter(function (f) { return !set[diaDe_(f[0])]; }).concat(nuevas);
  todas.sort(orden || function (x, y) { return x[0] - y[0]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, ncol).clearContent();
  if (todas.length) sh.getRange(2, 1, todas.length, ncol).setValues(todas);
}

function guardarVentas_(a) {
  if (!a.lista.length) throw new Error('El informe de ventas no trae ninguna venta con fecha.');
  var sp = hojaDatos_(TPV.PLATOS, TPV.PLATOS_COLS);
  var previos = leerDatos_(sp, 5);

  // Protección: si un mes tiene ventas por plato de antes (migradas, sin detalle por día) y todavía no hay
  // datos por día, hay que empezar con el informe completo («Este año») para no perder lo anterior.
  var T = L.TV, tv = libro_().getSheetByName(L.TPV), n = T.last - T.first + 1;
  var vt = tv.getRange(T.first, 1, n, 7).getValues();
  var conDia = {}; previos.forEach(function (f) { conDia[mesDeCelda_(f[0])] = true; });
  a.meses.forEach(function (m) {
    var legado = vt.some(function (f) { return f[1] !== '' && mesDeCelda_(f[0]) === m && f[2] !== TPV.ORIGEN; });
    var dias = a.lista.filter(function (d) { return d.slice(0, 7) === m; });
    var completo = a.meses.length > 1 || +dias[0].slice(8) <= 3;
    if (legado && !conDia[m] && !completo)
      throw new Error('En ' + m + ' hay ventas por plato de antes sin detalle diario. Sube primero el informe de ventas de «Este año» (o del mes completo); después ya podrás subir días sueltos.');
  });

  // 1) Platos por día
  var nuevasP = Object.keys(a.prod).map(function (k) { var p = a.prod[k]; return [aFecha_(p.dia), p.nombre, r2_(p.u), r2_(p.ul), r2_(p.tot)]; });
  reemplazarDias_(sp, 5, a.lista, nuevasP);
  // 2) Ventas TPV (resumen mensual por plato) de los meses tocados, reconstruido desde los días
  var set = {}; a.meses.forEach(function (m) { set[m] = true; });
  var agg = {};
  leerDatos_(sp, 5).forEach(function (f) {
    var m = mesDeCelda_(f[0]); if (!set[m]) return;
    var k = m + '|' + f[1], x = agg[k] || (agg[k] = [aFecha_(m + '-01'), f[1], TPV.ORIGEN, 0, 0, 0]);
    x[3] += f[2]; x[5] += f[3]; x[4] += f[4];
  });
  var quedan = vt.filter(function (f) { return f[1] !== '' && !set[mesDeCelda_(f[0])]; }).map(function (f) { return [f[0], f[1], f[2], f[4], f[5], f[6]]; });
  var nuevas = Object.keys(agg).map(function (k) { var x = agg[k]; return [x[0], x[1], x[2], r2_(x[3]), r2_(x[4]), r2_(x[5])]; });
  var todas = quedan.concat(nuevas);
  if (todas.length > n) throw new Error('No caben las ventas por plato (' + todas.length + ' filas, máximo ' + n + ').');
  tv.getRangeList(['A' + T.first + ':C' + T.last, 'E' + T.first + ':G' + T.last]).clearContent();
  if (todas.length) {
    tv.getRange(T.first, 1, todas.length, 3).setValues(todas.map(function (f) { return f.slice(0, 3); }));
    tv.getRange(T.first, 5, todas.length, 3).setValues(todas.map(function (f) { return f.slice(3, 6); }));
  }
  // 3) TPV Días (columnas de ventas; conserva las de tickets)
  var sd = hojaDatos_(TPV.DIAS, TPV.DIAS_COLS), nc = TPV.DIAS_COLS.length, prev = {};
  leerDatos_(sd, nc).forEach(function (f) { prev[diaDe_(f[0])] = f; });
  var dNuevas = a.lista.map(function (dia) {
    var d = a.dias[dia], p = prev[dia] || [];
    return [aFecha_(dia), dow_(dia), r2_(d.e), d.p, r2_(d.can.Sala), r2_(d.can.Terraza), r2_(d.can.Llevar), r2_(d.can.Recoger),
            r2_(d.can.Domicilio), d.pl, r2_(d.canc), p[11] || '', p[12] || '', p[13] || '', p[14] || '', p[15] || ''];
  });
  reemplazarDias_(sd, nc, a.lista, dNuevas);
  // 4) Horas y camareros por día
  var sh = hojaDatos_(TPV.HORAS, TPV.HORAS_COLS);
  reemplazarDias_(sh, 5, a.lista, Object.keys(a.horas).map(function (k) {
    var p = k.split('|'), h = a.horas[k]; return [aFecha_(p[0]), dow_(p[0]), +p[1], r2_(h.e), h.p];
  }), function (x, y) { return (x[0] - y[0]) || x[2] - y[2]; });
  var sc = hojaDatos_(TPV.CAM, TPV.CAM_COLS);
  reemplazarDias_(sc, 6, a.lista, Object.keys(a.cams).map(function (k) {
    var c = a.cams[k]; return [aFecha_(c.dia), c.cam, r2_(c.e), c.p, r2_(c.ce), r2_(c.cu)];
  }));
  SpreadsheetApp.flush();

  // Si ya había tickets de esos días, recoloca los ingresos por canal con el reparto real
  var ing = rellenarIngresos_(a.lista);
  var sinPlato = tv.getRange(T.first, 4, Math.max(todas.length, 1), 1).getValues().filter(function (f) { return f[0] === 'Sin asignar'; }).length;
  return { tipo: 'ventas', dias: a.lista.length, desde: a.lista[0], hasta: a.lista[a.lista.length - 1], sinAsignar: sinPlato, ingresos: ing };
}

function guardarFacturas_(a) {
  if (!a.lista.length) throw new Error('El informe de facturas no trae tickets con fecha.');
  var sd = hojaDatos_(TPV.DIAS, TPV.DIAS_COLS), nc = TPV.DIAS_COLS.length, prev = {};
  leerDatos_(sd, nc).forEach(function (f) { prev[diaDe_(f[0])] = f; });
  var nuevas = a.lista.map(function (dia) {
    var p = prev[dia] || [aFecha_(dia), dow_(dia), '', '', '', '', '', '', '', '', ''], t = a.dias[dia];
    var fila = p.slice(0, 11); fila[0] = aFecha_(dia); fila[1] = dow_(dia);
    return fila.concat([t.n, t.n ? r2_(t.tot / t.n) : '', r2_(t.ef), r2_(t.ta), r2_(t.ot)]);
  });
  reemplazarDias_(sd, nc, a.lista, nuevas);
  SpreadsheetApp.flush();
  var n = 0; a.lista.forEach(function (k) { n += a.dias[k].n; });
  var ing = rellenarIngresos_(a.lista);
  return { tipo: 'facturas', tickets: n, dias: a.lista.length, desde: a.lista[0], hasta: a.lista[a.lista.length - 1], ingresos: ing };
}

/* ------------------------------------------------ ingresos desde el TPV */

/**
 * Crea o actualiza las líneas de ingresos de cada día a partir de TPV Días (cobros del informe de facturas
 * y reparto por canal del informe de ventas). No toca días con ingresos apuntados a mano
 * (salvo los de plataformas de reparto, que no pasan por el TPV).
 */
function rellenarIngresos_(dias) {
  if (!ingresosAutomaticos_()) return { dias: 0, saltados: [], desactivado: true };
  var ss = libro_(), cfg = cfg_(), I = L.ING, n = I.last - I.first + 1;
  var tipoCanal = {}; cfg.canales.forEach(function (c) { tipoCanal[c.canal] = c.tipo; });
  var existe = function (c) { return tipoCanal.hasOwnProperty(c); };
  var canalSala = existe('Sala') ? 'Sala' : (cfg.canales[0] || {}).canal || '';
  var canalLlevar = existe('Para llevar') ? 'Para llevar' : canalSala;
  var canalReparto = existe('Reparto propio') ? 'Reparto propio' : canalLlevar;
  var catVenta = cfg.catIngreso[0] || '';

  var tpv = {}; leerDatos_(hojaDatos_(TPV.DIAS, TPV.DIAS_COLS), TPV.DIAS_COLS.length).forEach(function (f) { tpv[diaDe_(f[0])] = f; });
  var res = { dias: 0, saltados: [] };
  var porMes = {};
  dias.forEach(function (d) {
    if (+d.slice(0, 4) !== cfg.anio) return;
    var f = tpv[d]; if (!f || f[11] === '' || !f[11]) return;          // sin tickets de ese día: no hay cobros
    (porMes[+d.slice(5, 7) - 1] = porMes[+d.slice(5, 7) - 1] || []).push(d);
  });
  Object.keys(porMes).forEach(function (mi) {
    var sh = ss.getSheetByName(L.MESES[mi]);
    var v = sh.getRange(I.first, 1, n, 15).getValues();
    var esAuto = function (f) { return String(f[14]).indexOf(TPV.MARCA) >= 0; };
    var manual = function (f) { return f[5] !== '' && !esAuto(f) && tipoCanal[f[3]] !== 'Plataforma'; };
    var manualSinDia = v.some(function (f) { return manual(f) && f[0] === ''; });
    porMes[mi].forEach(function (d) {
      var dia = +d.slice(8);
      if (manualSinDia) { res.saltados.push(d + ' (el mes tiene ingresos a mano sin día)'); return; }
      if (v.some(function (f) { return manual(f) && +f[0] === dia; })) { res.saltados.push(d + ' (ya tiene ingresos a mano)'); return; }
      // Quita las líneas automáticas anteriores de ese día
      v.forEach(function (f, i) { if (esAuto(f) && +f[0] === dia) v[i] = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '']; });
      var t = tpv[d], cobrado = (t[13] || 0) + (t[14] || 0) + (t[15] || 0);
      if (!cobrado) return;
      // Reparto por canal según las ventas del TPV de ese día (si no hay informe de ventas, todo a sala)
      var sala = (t[4] || 0) + (t[5] || 0), llevar = (t[6] || 0) + (t[7] || 0), reparto = t[8] || 0, tot = sala + llevar + reparto;
      var partes = tot ? [[canalSala, sala / tot], [canalLlevar, llevar / tot], [canalReparto, reparto / tot]] : [[canalSala, 1]];
      var junt = {}; partes.forEach(function (p) { if (p[1] > 0) junt[p[0]] = (junt[p[0]] || 0) + p[1]; });
      var canales = Object.keys(junt).sort(function (x, y) { return junt[y] - junt[x]; });
      var reparte = function (total) {
        var acc = 0; return canales.map(function (c, k) {
          var x = k === canales.length - 1 ? r2_(total - acc) : r2_(total * junt[c]); acc = r2_(acc + x); return x;
        });
      };
      var brutos = reparte(cobrado), tarjetas = reparte(t[14] || 0), otros = reparte(t[15] || 0);
      canales.forEach(function (c, k) {
        var libre = v.findIndex(function (f) { return f.join('') === ''; });
        if (libre < 0) throw new Error('El registro de ingresos de ' + L.MESES[mi] + ' está lleno.');
        v[libre] = [dia, '', 'Cierre TPV', c, catVenta, brutos[k], tarjetas[k] || '', otros[k] || '', '', '', '', '', k === 0 ? t[11] : '', '',
                    TPV.MARCA + (k === 0 ? ' · ' + t[11] + ' tickets' : '')];
      });
      res.dias++;
    });
    // Escribe solo las columnas que rellena el usuario (A, C:H, O); las demás son fórmulas
    sh.getRange(I.first, 1, n, 1).setValues(v.map(function (f) { return [f[0]]; }));
    sh.getRange(I.first, 3, n, 6).setValues(v.map(function (f) { return f.slice(2, 8); }));
    sh.getRange(I.first, 13, n, 3).setValues(v.map(function (f) { return f.slice(12, 15); }));
  });
  SpreadsheetApp.flush();
  return res;
}

/** Asigna nombres del TPV a platos (tabla de la derecha en Ventas TPV). pares = [[nombreTPV, plato], ...] */
function asignarNombresTpv(pares) {
  var sh = libro_().getSheetByName(L.TPV), M = L.MAP;
  var v = sh.getRange(M.first, 9, M.last - M.first + 1, 2).getValues();
  var idx = {}; v.forEach(function (f, i) { if (f[0] !== '') idx[String(f[0]).toLowerCase()] = i; });
  pares.forEach(function (p) {
    if (!p[0] || !p[1]) return;
    var i = idx[String(p[0]).toLowerCase()];
    if (i === undefined) { i = v.findIndex(function (f) { return f[0] === ''; }); if (i < 0) throw new Error('La tabla de nombres está llena.'); idx[String(p[0]).toLowerCase()] = i; }
    v[i] = [p[0], p[1]];
  });
  sh.getRange(M.first, 9, v.length, 2).setValues(v);
  SpreadsheetApp.flush();
  return true;
}

/* -------------------------------------------------------------- lectura */

function filasDatos_(nombre, ncol) {
  var sh = libro_().getSheetByName(nombre);
  return sh ? leerDatos_(sh, ncol) : [];
}

/** Coste de ración, envase, margen y categoría de cada plato de la carta. */
function mapaPlatos_() {
  var out = {};
  platos_().forEach(function (p) { out[p.plato] = p; });
  var sh = libro_().getSheetByName(L.PLATOS);
  sh.getRange(L.PL.first, 1, L.PL.last - L.PL.first + 1, 5).getValues().forEach(function (f) {
    if (f[0] !== '' && out[f[0]]) out[f[0]].envase = typeof f[4] === 'number' ? f[4] : 0;
  });
  return out;
}

/** Todo lo del TPV para un mes (nombre de mes de la plantilla). */
function ventasTpvMes_(mes) {
  var cfg = cfg_(), idx = L.MESES.indexOf(mes);
  var clave = cfg.anio + '-' + ('0' + (idx + 1)).slice(-2);
  var T = L.TV, sh = libro_().getSheetByName(L.TPV);
  var filas = sh.getRange(T.first, 1, T.last - T.first + 1, 7).getValues()
    .filter(function (f) { return f[1] !== '' && mesDeCelda_(f[0]) === clave; });
  var pl = mapaPlatos_();
  var enMes = function (f) { return mesDeCelda_(f[0]) === clave; };
  var dias = filasDatos_(TPV.DIAS, TPV.DIAS_COLS.length).filter(enMes);
  var horas = filasDatos_(TPV.HORAS, 5).filter(enMes);
  var camsD = filasDatos_(TPV.CAM, 6).filter(enMes);
  var sumD = function (i) { return dias.reduce(function (s, f) { return s + (typeof f[i] === 'number' ? f[i] : 0); }, 0); };

  // Por plato
  var porPlato = {}, sinAsignar = [], importeTotal = 0, importeCubierto = 0, teorico = 0, envases = 0;
  filas.forEach(function (f) {
    var nombre = f[3], u = f[4] || 0, imp = f[5] || 0, ul = f[6] || 0;
    importeTotal += imp;
    if (nombre === 'Sin asignar') { sinAsignar.push({ tpv: f[1], unidades: u, importe: r2_(imp) }); return; }
    var p = pl[nombre];
    var x = porPlato[nombre] || (porPlato[nombre] = { plato: nombre, categoria: p ? p.categoria : '', unidades: 0, llevar: 0, importe: 0 });
    x.unidades += u; x.llevar += ul; x.importe += imp;
    if (p && p.coste > 0) { importeCubierto += imp; teorico += u * p.coste + ul * (p.envase || 0); envases += ul * (p.envase || 0); }
  });
  var lista = Object.keys(porPlato).map(function (k) {
    var x = porPlato[k], p = pl[k] || {};
    x.importe = r2_(x.importe); x.margenUd = p.margen; x.margenTotal = p.margen != null ? r2_(p.margen * x.unidades) : null;
    x.foodCost = p.foodCostPct; x.clasificacion = p.clasificacion; return x;
  });

  // CMV teórico frente a real (en % sobre ventas, como el panel antiguo)
  var bolsa = Number(libro_().getSheetByName(L.CFG).getRange('B9').getValue()) || 0;
  var pedLlevar = sumD(9);
  var ventasSinIvaCubiertas = importeCubierto / (1 + (cfg.ivaVentas || 0.1));
  var rm = resumenMes_(mes);
  var cmv = null;
  if (ventasSinIvaCubiertas > 0) {
    var teoricoTot = teorico + pedLlevar * bolsa;
    var pctTeo = teoricoTot / ventasSinIvaCubiertas;
    var pctReal = rm.ventas ? rm.cmv / rm.ventas : null;
    cmv = { teoricoPct: pct_(pctTeo), realPct: pctReal != null ? pct_(pctReal) : null,
            cobertura: importeTotal ? pct_(importeCubierto / importeTotal) : null,
            descuadre: pctReal != null ? r2_((pctReal - pctTeo) * rm.ventas) : null,
            envases: r2_(envases), bolsas: r2_(pedLlevar * bolsa), costeBolsa: bolsa,
            nota: rm.cmv === rm.compras ? 'Sin inventario: el CMV real son las compras del mes.' : '' };
  }

  // Día de la semana, horas y camareros (agregados del mes)
  var porDow = TPV.DOW.map(function (d) {
    var fs = dias.filter(function (f) { return f[1] === d && f[2] > 0; });
    var tot = fs.reduce(function (s, f) { return s + f[2]; }, 0);
    return { dia: d, noches: fs.length, media: fs.length ? r2_(tot / fs.length) : 0 };
  });
  var hAgg = {};
  horas.forEach(function (f) { var k = f[1] + '|' + f[2], x = hAgg[k] || (hAgg[k] = { dia: f[1], hora: f[2], ventas: 0, pedidos: 0 }); x.ventas += f[3]; x.pedidos += f[4]; });
  var cAgg = {};
  camsD.forEach(function (f) { var x = cAgg[f[1]] || (cAgg[f[1]] = { camarero: f[1], ventas: 0, pedidos: 0, cancelado: 0 }); x.ventas += f[2]; x.pedidos += f[3]; x.cancelado += f[4]; });
  var tickets = sumD(11), cobrado = sumD(13) + sumD(14) + sumD(15);
  return {
    mes: mes, conDatos: dias.length > 0 || filas.length > 0, ingAuto: ingresosAutomaticos_(),
    desde: dias.length ? iso_(dias[0][0]) : null, hasta: dias.length ? iso_(dias[dias.length - 1][0]) : null,
    noches: dias.filter(function (f) { return f[2] > 0; }).length,
    ventas: r2_(sumD(2)), pedidos: sumD(3), pedidosLlevar: pedLlevar, cancelado: r2_(sumD(10)),
    canales: { Sala: r2_(sumD(4)), Terraza: r2_(sumD(5)), 'Para llevar': r2_(sumD(6)), Recoger: r2_(sumD(7)), Domicilio: r2_(sumD(8)) },
    tickets: tickets, ticketMedio: tickets ? r2_(cobrado / tickets) : null,
    cobros: { efectivo: r2_(sumD(13)), tarjeta: r2_(sumD(14)), otros: r2_(sumD(15)) },
    platos: lista, sinAsignar: sinAsignar, cmv: cmv, porDiaSemana: porDow,
    horas: Object.keys(hAgg).map(function (k) { var x = hAgg[k]; x.ventas = r2_(x.ventas); return x; }),
    camareros: Object.keys(cAgg).map(function (k) {
      var x = cAgg[k], base = x.ventas + x.cancelado;
      return { camarero: x.camarero, ventas: r2_(x.ventas), pedidos: x.pedidos, cancelado: r2_(x.cancelado), pctCancelado: base ? Math.round(x.cancelado / base * 1000) / 10 : 0 };
    }).sort(function (a, b) { return b.ventas - a.ventas; })
  };
}
