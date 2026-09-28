/**
 * Lectura de datos de la plantilla para el panel y el asistente.
 * Todos los cálculos los hace la hoja; aquí solo se leen y se ordenan.
 */

function r2_(v) { return typeof v === 'number' ? Math.round(v * 100) / 100 : (v === '' ? null : v); }
function pct_(v) { return typeof v === 'number' ? Math.round(v * 1000) / 10 : null; }

function cfg_() {
  var sh = libro_().getSheetByName(L.CFG);
  var v = sh.getRange('A1:P' + L.C.cat.last).getValues();
  var col = function (c, last, w) {
    var i = c.charCodeAt(0) - 65, out = [];
    for (var r = L.C.listFirst - 1; r < last; r++) {
      if (v[r][i] === '') continue;
      out.push(w ? v[r].slice(i, i + w) : v[r][i]);
    }
    return out;
  };
  return {
    nombre: v[3][1] || 'tu restaurante', anio: Number(v[4][1]) || new Date().getFullYear(),
    ivaVentas: v[7][1],
    umbrales: { foodCost: [v[11][1], v[11][2]], personal: [v[12][1], v[12][2]], prime: [v[13][1], v[13][2]],
                margen: [v[14][1], v[14][2]], foodCostPlato: [v[15][1], v[15][2]] },
    categorias: col('A', L.C.cat.last, 2).map(function (x) { return { categoria: x[0], tipo: x[1] }; }),
    canales: col('H', L.C.canal.last, 2).map(function (x) { return { canal: x[0], tipo: x[1] }; }),
    catIngreso: col('K', L.C.catIng.last),
    formasPago: col('N', L.C.pago.last)
  };
}

/** Lee el bloque de resumen de un mes (una sola llamada a la hoja). */
function resumenMes_(mes) {
  var sh = libro_().getSheetByName(mes);
  if (!sh) throw new Error('No existe la hoja "' + mes + '".');
  var v = sh.getRange('A1:O35').getValues();
  var at = function (a1) { var m = a1.match(/^([A-Z])(\d+)$/); return v[+m[2] - 1][m[1].charCodeAt(0) - 65]; };
  var M = L.M, out = { mes: mes };
  ['ventas', 'compras', 'varInv', 'cmv', 'margenBruto', 'personal', 'explotacion', 'resultado', 'inversiones',
   'financiacion', 'impuestos', 'retirada', 'cajaGenerada', 'saldoIni', 'cobrado', 'pagado', 'saldoPrev',
   'saldoReal', 'diferencia', 'pendiente', 'cobEfectivo', 'cobTarjeta', 'cobBizum', 'cobPlataformas',
   'invIni', 'invFin', 'ticketMedio', 'gastoComensal', 'ventaDia', 'costesFijos', 'ventaNecesaria',
   'ventaNecDia', 'difEquilibrio'].forEach(function (k) { out[k] = r2_(at(M[k])); });
  out.dias = at(M.dias) || 0; out.tickets = r2_(at(M.tickets)); out.comensales = r2_(at(M.comensales));
  out.kpi = { foodCost: pct_(at(M.kFoodCost)), personal: pct_(at(M.kPersonal)), prime: pct_(at(M.kPrime)),
              margen: pct_(at(M.kMargen)), explotacion: pct_(at(M.kExplot)) };
  out.estado = { foodCost: at('E25'), personal: at('E26'), prime: at('E27'), margen: at('E28') };
  out.margenContrib = pct_(at(M.margenContrib));
  out.gastosPorCategoria = [];
  for (var r = M.catFirst; r <= M.catLast; r++) {
    var fila = v[r - 1];
    if (fila[5] !== '' && typeof fila[7] === 'number' && fila[7] !== 0)
      out.gastosPorCategoria.push({ categoria: fila[5], tipo: fila[6], importe: r2_(fila[7]), pctVentas: pct_(fila[8]) });
  }
  out.gastosSinCategoria = r2_(v[34][7]);
  out.conDatos = (out.ventas || 0) !== 0 || out.gastosPorCategoria.length > 0;
  return out;
}

function resumenAnual_() {
  var meses = L.MESES.map(resumenMes_);
  var sum = function (k) { return r2_(meses.reduce(function (s, m) { return s + (m[k] || 0); }, 0)); };
  var ventas = sum('ventas');
  var tot = { ventas: ventas, cmv: sum('cmv'), personal: sum('personal'), explotacion: sum('explotacion'),
              resultado: sum('resultado'), inversiones: sum('inversiones'), financiacion: sum('financiacion'),
              impuestos: sum('impuestos'), retirada: sum('retirada'), cajaGenerada: sum('cajaGenerada'),
              cobrado: sum('cobrado'), pagado: sum('pagado') };
  if (ventas) {
    tot.foodCostPct = pct_(tot.cmv / ventas); tot.personalPct = pct_(tot.personal / ventas);
    tot.primePct = pct_((tot.cmv + tot.personal) / ventas); tot.margenPct = pct_(tot.resultado / ventas);
  }
  return { meses: meses, total: tot };
}

/** Filas de gastos. mes: nombre, o 'todos'. */
function gastos_(mes) {
  var G = L.GAS, n = G.last - G.first + 1, out = [];
  (mes && mes !== 'todos' ? [mes] : L.MESES).forEach(function (m) {
    var sh = libro_().getSheetByName(m);
    sh.getRange(G.first, 1, n, 13).getValues().forEach(function (f, i) {
      if (f[3] === '' && f[4] === '' && f[5] === '' && f[7] === '') return;
      out.push({ mes: m, fila: G.first + i, dia: f[0], factura: f[2], proveedor: f[3], concepto: f[4],
                 categoria: f[5], tipo: f[6], base: r2_(f[7]), iva: r2_(f[8]), total: r2_(f[9]), pago: f[10], obs: f[12] });
    });
  });
  return out;
}

function ingresos_(mes) {
  var I = L.ING, n = I.last - I.first + 1, out = [];
  (mes && mes !== 'todos' ? [mes] : L.MESES).forEach(function (m) {
    var sh = libro_().getSheetByName(m);
    sh.getRange(I.first, 1, n, 15).getValues().forEach(function (f, i) {
      if (f[5] === '') return;
      out.push({ mes: m, dia: f[0], concepto: f[2], canal: f[3], categoria: f[4], bruto: r2_(f[5]), tarjeta: r2_(f[6]),
                 bizum: r2_(f[7]), efectivo: r2_(f[8]), plataformas: r2_(f[9]), sinIva: r2_(f[10]),
                 tickets: f[12] || null, comensales: f[13] || null, obs: f[14] });
    });
  });
  return out;
}

function platos_() {
  var P = L.PL, sh = libro_().getSheetByName(L.PLATOS);
  return sh.getRange(P.first, 1, P.last - P.first + 1, 14).getValues().filter(function (f) { return f[0] !== ''; })
    .map(function (f) {
      return { plato: f[0], categoria: f[1], pvp: r2_(f[2]), coste: r2_(f[5]), pvpSinIva: r2_(f[6]), foodCostPct: pct_(f[7]),
               margen: r2_(f[8]), foodCostDeliveryPct: pct_(f[9]), estado: f[10], unidades: f[11] || 0,
               margenTotal: r2_(f[12]), clasificacion: f[13] };
    });
}

function agrupar_(filas, clave, campo) {
  var acc = {};
  filas.forEach(function (f) {
    var k = f[clave] === '' || f[clave] == null ? '(sin dato)' : String(f[clave]);
    acc[k] = acc[k] || { grupo: k, importe: 0, lineas: 0 };
    acc[k].importe += f[campo] || 0; acc[k].lineas++;
  });
  return Object.keys(acc).map(function (k) { acc[k].importe = r2_(acc[k].importe); return acc[k]; })
    .sort(function (a, b) { return b.importe - a.importe; });
}

/** Mes con datos más reciente (para abrir el panel en él). */
function ultimoMesConDatos_() {
  var hoy = new Date(), anio = cfg_().anio;
  var desde = anio === hoy.getFullYear() ? hoy.getMonth() : 11;
  for (var i = desde; i >= 0; i--) { if (resumenMes_(L.MESES[i]).conDatos) return L.MESES[i]; }
  return L.MESES[desde];
}
