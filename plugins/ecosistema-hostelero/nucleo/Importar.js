/**
 * Cargar datos preparados por Claude (plugin «Ecosistema Hostelero»).
 * El Claude del hostelero lee sus Excel, su carta o sus fichas y deja un archivo PersonalDatos.js con
 * `var DATOS_CLAUDE = {...}`. El dueño pulsa 🍽️ Ecosistema › 📥 Cargar datos preparados por Claude,
 * ve un resumen y confirma. Solo se escriben las columnas que rellena el usuario (las fórmulas no se tocan)
 * y siempre en líneas libres; nada se borra salvo las líneas de receta de los platos que se vuelven a cargar.
 *
 * Formato (todo opcional salvo `id`):
 * {
 *   id: 'migracion-2026-01',                // para no cargar dos veces lo mismo
 *   descripcion: 'Mis cuentas de enero a septiembre',
 *   configuracion: { nombre, anio, saldoInicial, inventarioInicial, ivaVentas, bolsaPorPedido, tpvDesde: 'AAAA-MM-DD' },
 *   categoriasGasto: [['Nombre', 'Tipo', 'Qué incluye'], …],   // se añaden las que falten
 *   canales: [['Nombre', 'Directo'|'Plataforma'], …],
 *   categoriasIngreso: [['Nombre', 0.10], …],
 *   formasPago: ['Tarjeta', …],
 *   categoriasCarta: ['Hamburguesas', …],
 *   locales: [['Nombre', 'enlace de su hoja'], …],
 *   ingresos: [{ mes: 1-12, dia, concepto, canal, categoria, bruto, tarjeta, bizum, tickets, comensales, observaciones }],
 *   gastos:   [{ mes: 1-12, dia, factura, proveedor, concepto, categoria, base, iva, formaPago, observaciones }],
 *   cierres:  [{ mes: 1-12, efectivo, bancos, inventarioFinal }],
 *   ingredientes: [{ nombre, unidad: 'kg'|'l'|'ud', proveedor, precio, cantidad, merma, notas }],  // si existe, se actualiza
 *   platos:   [{ nombre, categoria, pvp, iva, envase }],                                        // si existe, se actualiza
 *   recetas:  [{ plato, ingrediente, cantidad }],   // sustituye la receta entera de cada plato que aparece
 *   ventasPlatos: [{ mes: 1-12, nombre, unidades, importe, udsLlevar, plato }]
 *     // ventas por plato de un TPV que no es Foodyservice (resumen mensual). `nombre` = como sale en el TPV;
 *     // `plato` (opcional) = plato de la carta si el nombre no coincide. Sustituye lo cargado antes así para ese mes.
 *   nombresTpv: [['NOMBRE EN EL TPV', 'Plato de la carta'], …]   // relaciona nombres del TPV con platos (tabla de Ventas TPV)
 * }
 */

function cargarDatosClaude() {
  var ui = SpreadsheetApp.getUi();
  if (typeof DATOS_CLAUDE === 'undefined' || !DATOS_CLAUDE) {
    ui.alert('Cargar datos', 'No hay datos preparados. Pídele a Claude que te los prepare («pasa mis Excel a la hoja») y vuelve a pulsar aquí.', ui.ButtonSet.OK);
    return;
  }
  var d = DATOS_CLAUDE, props = PropertiesService.getDocumentProperties();
  var plan;
  try { plan = planCarga_(d); } catch (e) { ui.alert('Cargar datos', 'No se ha cargado nada: ' + e.message, ui.ButtonSet.OK); return; }
  var yaCargados = JSON.parse(props.getProperty('DATOS_CLAUDE_CARGADOS') || '[]');
  var aviso = yaCargados.indexOf(d.id) >= 0 ? '⚠️ ESTOS DATOS YA SE CARGARON UNA VEZ. Si sigues, se duplicarán ingresos y gastos.\n\n' : '';
  var ok = ui.alert('Cargar datos preparados por Claude',
    aviso + (d.descripcion ? d.descripcion + '\n\n' : '') + plan.resumen.join('\n') +
    (plan.avisos.length ? '\n\nAvisos:\n' + plan.avisos.join('\n') : '') + '\n\n¿Cargar ahora?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var hecho = ejecutarCarga_(d);
  yaCargados.push(d.id);
  props.setProperty('DATOS_CLAUDE_CARGADOS', JSON.stringify(yaCargados));
  ui.alert('Cargar datos', 'Hecho.\n\n' + hecho.join('\n') + '\n\nYa puedes decirle a Claude que retire el archivo de datos.', ui.ButtonSet.OK);
}

/* --------------------------------------------------------- comprobar */

function mesIdx_(m) {
  var i = typeof m === 'number' ? m - 1 : L.MESES.indexOf(String(m));
  if (!(i >= 0 && i < 12)) throw new Error('Mes no válido: «' + m + '» (usa 1 a 12).');
  return i;
}

/** Comprueba todo ANTES de escribir nada. Devuelve el resumen para el dueño. */
function planCarga_(d) {
  if (!d.id) throw new Error('A los datos les falta el identificador (id).');
  var cfg = cfg_(), resumen = [], avisos = [];
  var conNombre = function (lista, extra) {
    return lista.concat((extra || []).map(function (x) { return Array.isArray(x) ? x[0] : x; }));
  };
  var cats = conNombre(cfg.categorias.map(function (c) { return c.categoria; }), d.categoriasGasto);
  var canales = conNombre(cfg.canales.map(function (c) { return c.canal; }), d.canales);
  var catIng = conNombre(cfg.catIngreso, d.categoriasIngreso);
  var pagos = conNombre(cfg.formasPago, d.formasPago);
  var falta = function (lista, valor, que, donde) {
    if (valor !== undefined && valor !== '' && lista.indexOf(valor) < 0) throw new Error(que + ' «' + valor + '» no existe en Configuración (' + donde + '). Añádela en los datos o cámbiala.');
  };
  var num = function (v, que) { if (v !== undefined && v !== '' && typeof v !== 'number') throw new Error(que + ' no es un número: «' + v + '».'); };

  if (d.configuracion) {
    var c = d.configuracion;
    if (c.anio && Number(c.anio) !== cfg.anio) avisos.push('• El año de la hoja pasa de ' + cfg.anio + ' a ' + c.anio + '.');
    resumen.push('• Datos del negocio' + (c.nombre ? ': ' + c.nombre : ''));
  }
  [['categoriasGasto', 'categorías de gasto'], ['canales', 'canales'], ['categoriasIngreso', 'categorías de ingreso'],
   ['formasPago', 'formas de pago'], ['categoriasCarta', 'categorías de la carta'], ['locales', 'locales del grupo']].forEach(function (x) {
    if (d[x[0]] && d[x[0]].length) resumen.push('• ' + d[x[0]].length + ' ' + x[1] + ' (se añaden las que falten)');
  });

  var porMes = function (filas, que) {
    var n = {}; filas.forEach(function (f) { var i = mesIdx_(f.mes); n[i] = (n[i] || 0) + 1; });
    return Object.keys(n).map(function (i) { return L.MESES[i] + ' ' + n[i]; }).join(', ');
  };
  if (d.ingresos && d.ingresos.length) {
    d.ingresos.forEach(function (f) {
      falta(canales, f.canal, 'El canal', 'canales'); falta(catIng, f.categoria, 'La categoría de ingreso', 'categorías de ingreso');
      num(f.bruto, 'Un importe de ingreso'); num(f.tarjeta, 'Un cobro con tarjeta'); num(f.bizum, 'Un cobro por Bizum');
    });
    resumen.push('• ' + d.ingresos.length + ' líneas de ingresos (' + porMes(d.ingresos) + ')');
  }
  if (d.gastos && d.gastos.length) {
    d.gastos.forEach(function (f) {
      falta(cats, f.categoria, 'La categoría de gasto', 'categorías de gasto'); falta(pagos, f.formaPago, 'La forma de pago', 'formas de pago');
      num(f.base, 'Una base de gasto'); num(f.iva, 'Un IVA de gasto');
      if (!f.categoria) avisos.push('• Hay gastos sin categoría: saldrán como «Revisar».');
    });
    resumen.push('• ' + d.gastos.length + ' líneas de gastos (' + porMes(d.gastos) + ')');
  }
  if (d.cierres && d.cierres.length) { d.cierres.forEach(function (f) { mesIdx_(f.mes); }); resumen.push('• Caja e inventario de ' + d.cierres.length + ' meses'); }

  var ingNombres = leerColumna_(L.INGRED, L.IN).concat((d.ingredientes || []).map(function (i) { return i.nombre; }));
  var plNombres = leerColumna_(L.PLATOS, L.PL).concat((d.platos || []).map(function (p) { return p.nombre; }));
  if (d.ingredientes && d.ingredientes.length) {
    d.ingredientes.forEach(function (i) {
      if (!i.nombre) throw new Error('Hay un ingrediente sin nombre.');
      if (['kg', 'l', 'ud'].indexOf(i.unidad) < 0) throw new Error('Unidad no válida en «' + i.nombre + '»: usa kg, l o ud.');
      if (i.merma !== undefined && i.merma !== '' && !(i.merma >= 0 && i.merma < 1)) throw new Error('Merma no válida en «' + i.nombre + '»: va de 0 a 0,99 (20 % = 0,2).');
    });
    resumen.push('• ' + d.ingredientes.length + ' ingredientes (nuevos o actualizados)');
  }
  if (d.platos && d.platos.length) {
    var carta = cfg_carta_().concat(d.categoriasCarta || []);
    d.platos.forEach(function (p) {
      if (!p.nombre) throw new Error('Hay un plato sin nombre.');
      falta(carta, p.categoria, 'La categoría de la carta', 'categorías de la carta'); num(p.pvp, 'El precio de «' + p.nombre + '»');
    });
    resumen.push('• ' + d.platos.length + ' platos (nuevos o actualizados)');
  }
  if (d.recetas && d.recetas.length) {
    var platosRec = {};
    d.recetas.forEach(function (r) {
      if (plNombres.indexOf(r.plato) < 0) throw new Error('La receta es de un plato que no existe: «' + r.plato + '».');
      if (ingNombres.indexOf(r.ingrediente) < 0) throw new Error('La receta de «' + r.plato + '» usa un ingrediente que no existe: «' + r.ingrediente + '».');
      num(r.cantidad, 'Una cantidad de «' + r.plato + '»');
      platosRec[r.plato] = 1;
    });
    resumen.push('• Recetas de ' + Object.keys(platosRec).length + ' platos (' + d.recetas.length + ' líneas; sustituyen a las que hubiera)');
  }
  if (d.ventasPlatos && d.ventasPlatos.length) {
    d.ventasPlatos.forEach(function (v) {
      if (!v.nombre) throw new Error('Hay una venta por plato sin nombre.');
      num(v.unidades, 'Las unidades de «' + v.nombre + '»'); num(v.importe, 'El importe de «' + v.nombre + '»');
      if (v.plato && plNombres.indexOf(v.plato) < 0) throw new Error('La venta de «' + v.nombre + '» apunta a un plato que no existe: «' + v.plato + '».');
    });
    var tv = libro_().getSheetByName(L.TPV).getRange(L.TV.first, 1, L.TV.last - L.TV.first + 1, 3).getValues();
    var mesesConTpv = {};
    tv.forEach(function (f) { if (f[1] !== '' && f[2] === TPV.ORIGEN && f[0] instanceof Date) mesesConTpv[f[0].getMonth()] = 1; });
    d.ventasPlatos.forEach(function (v) {
      var i = mesIdx_(v.mes);
      if (mesesConTpv[i]) avisos.push('• ' + L.MESES[i] + ' ya tiene ventas por plato de los informes del panel: se sumarán las dos.');
    });
    resumen.push('• ' + d.ventasPlatos.length + ' ventas por plato (' + porMes(d.ventasPlatos) + ')');
  }
  if (d.nombresTpv && d.nombresTpv.length) {
    d.nombresTpv.forEach(function (p) {
      if (!p[0] || !p[1]) throw new Error('Hay una relación de nombres del TPV incompleta.');
      if (plNombres.indexOf(p[1]) < 0) throw new Error('«' + p[0] + '» apunta a un plato que no existe: «' + p[1] + '».');
    });
    resumen.push('• ' + d.nombresTpv.length + ' nombres del TPV relacionados con platos');
  }
  if (!resumen.length) throw new Error('Los datos están vacíos.');
  return { resumen: resumen, avisos: avisos.filter(function (a, i, arr) { return arr.indexOf(a) === i; }) };
}

function cfg_carta_() {
  var sh = libro_().getSheetByName(L.CFG);
  return sh.getRange('P' + L.C.listFirst + ':P' + L.C.carta.last).getValues().map(function (f) { return f[0]; }).filter(String);
}
function leerColumna_(hoja, blk) {
  return libro_().getSheetByName(hoja).getRange('A' + blk.first + ':A' + blk.last).getValues().map(function (f) { return f[0]; }).filter(String);
}

/* ------------------------------------------------------------ escribir */

function ejecutarCarga_(d) {
  var ss = libro_(), hecho = [];
  var cfgSh = ss.getSheetByName(L.CFG);

  if (d.configuracion) {
    var c = d.configuracion, set = function (a1, v) { if (v !== undefined && v !== null && v !== '') cfgSh.getRange(a1).setValue(v); };
    set(L.C.nombre, c.nombre); set(L.C.anio, c.anio); set(L.C.saldoIni, c.saldoInicial); set(L.C.invIni, c.inventarioInicial);
    set(L.C.ivaVentas, c.ivaVentas); set('B9', c.bolsaPorPedido);
    if (/^\d{4}-\d{2}-\d{2}$/.test(c.tpvDesde || '')) { var p = c.tpvDesde.split('-'); set(L.C.tpvDesde, new Date(+p[0], p[1] - 1, +p[2])); }
    hecho.push('• Datos del negocio');
  }
  [['categoriasGasto', 'cat', 3, 'categorías de gasto'], ['canales', 'canal', 2, 'canales'], ['categoriasIngreso', 'catIng', 2, 'categorías de ingreso'],
   ['formasPago', 'pago', 1, 'formas de pago'], ['categoriasCarta', 'carta', 1, 'categorías de la carta'], ['locales', 'locales', 2, 'locales']].forEach(function (x) {
    if (!d[x[0]] || !d[x[0]].length) return;
    var n = anadirALista_(cfgSh, x[1], x[2], d[x[0]].map(function (v) { return Array.isArray(v) ? v : [v]; }));
    if (n) hecho.push('• ' + n + ' ' + x[3] + ' nuevas');
  });

  var agrupar = function (filas) { var g = {}; (filas || []).forEach(function (f) { var i = mesIdx_(f.mes); (g[i] = g[i] || []).push(f); }); return g; };
  var ing = agrupar(d.ingresos), gas = agrupar(d.gastos);
  Object.keys(ing).forEach(function (i) {
    // A | C:H | M:O  (B, I:L son fórmulas)
    escribirFilas_(ss.getSheetByName(L.MESES[i]), L.ING, [[1, 1], [3, 6], [13, 3]], ing[i].map(function (f) {
      return [[v_(f.dia)], [v_(f.concepto), v_(f.canal), v_(f.categoria), v_(f.bruto), v_(f.tarjeta), v_(f.bizum)],
              [v_(f.tickets), v_(f.comensales), v_(f.observaciones)]];
    }), 'ingresos de ' + L.MESES[i]);
  });
  if (d.ingresos && d.ingresos.length) hecho.push('• ' + d.ingresos.length + ' líneas de ingresos');
  Object.keys(gas).forEach(function (i) {
    // A | C:F | H:I | K | M  (B, G, J son fórmulas; L es el enlace a la factura)
    escribirFilas_(ss.getSheetByName(L.MESES[i]), L.GAS, [[1, 1], [3, 4], [8, 2], [11, 1], [13, 1]], gas[i].map(function (f) {
      return [[v_(f.dia)], [v_(f.factura), v_(f.proveedor), v_(f.concepto), v_(f.categoria)], [r2v_(f.base), r2v_(f.iva)],
              [v_(f.formaPago)], [v_(f.observaciones)]];
    }), 'gastos de ' + L.MESES[i]);
  });
  if (d.gastos && d.gastos.length) hecho.push('• ' + d.gastos.length + ' líneas de gastos');
  (d.cierres || []).forEach(function (f) {
    var sh = ss.getSheetByName(L.MESES[mesIdx_(f.mes)]);
    if (f.efectivo !== undefined) sh.getRange(L.M.efectivoReal).setValue(f.efectivo);
    if (f.bancos !== undefined) sh.getRange(L.M.bancosReal).setValue(f.bancos);
    if (f.inventarioFinal !== undefined) sh.getRange(L.M.invFin).setValue(f.inventarioFinal);
  });
  if (d.cierres && d.cierres.length) hecho.push('• Caja e inventario de ' + d.cierres.length + ' meses');

  if (d.ingredientes && d.ingredientes.length) {
    // A:E | G | I:J  (F y H son fórmulas)
    var hoy = new Date();
    actualizarPorNombre_(ss.getSheetByName(L.INGRED), L.IN, [[1, 5], [7, 1], [9, 2]], d.ingredientes.map(function (i) {
      return [[i.nombre, i.unidad, v_(i.proveedor), v_(i.precio), v_(i.cantidad)], [v_(i.merma)], [hoy, v_(i.notas)]];
    }), 'ingredientes');
    hecho.push('• ' + d.ingredientes.length + ' ingredientes');
  }
  if (d.platos && d.platos.length) {
    actualizarPorNombre_(ss.getSheetByName(L.PLATOS), L.PL, [[1, 5]], d.platos.map(function (p) {
      return [[p.nombre, v_(p.categoria), v_(p.pvp), v_(p.iva), v_(p.envase)]];
    }), 'platos');
    hecho.push('• ' + d.platos.length + ' platos');
  }
  if (d.recetas && d.recetas.length) {
    var sh = ss.getSheetByName(L.RECETAS), R = L.RE, rango = sh.getRange(R.first, 1, R.last - R.first + 1, 3);
    var v = rango.getValues(), platos = {};
    d.recetas.forEach(function (r) { platos[r.plato] = 1; });
    v = v.map(function (f) { return platos[f[0]] ? ['', '', ''] : f; });   // fuera la receta anterior de esos platos
    d.recetas.forEach(function (r) {
      var libre = v.findIndex(function (f) { return f.join('') === ''; });
      if (libre < 0) throw new Error('La hoja de Recetas está llena.');
      v[libre] = [r.plato, r.ingrediente, r.cantidad];
    });
    rango.setValues(v);
    hecho.push('• Recetas de ' + Object.keys(platos).length + ' platos');
  }
  if (d.ventasPlatos && d.ventasPlatos.length) {
    var anio = cfg_().anio, T = L.TV, tvSh = ss.getSheetByName(L.TPV), nT = T.last - T.first + 1;
    var meses = {}; d.ventasPlatos.forEach(function (v) { meses[mesIdx_(v.mes)] = 1; });
    // Fuera lo cargado antes por Claude en esos meses (A:C y E:G; D es fórmula)
    var ac = tvSh.getRange(T.first, 1, nT, 3).getValues(), eg = tvSh.getRange(T.first, 5, nT, 3).getValues();
    ac.forEach(function (f, i) {
      if (f[2] === ORIGEN_CLAUDE && f[0] instanceof Date && f[0].getFullYear() === anio && meses[f[0].getMonth()]) { ac[i] = ['', '', '']; eg[i] = ['', '', '']; }
    });
    tvSh.getRange(T.first, 1, nT, 3).setValues(ac); tvSh.getRange(T.first, 5, nT, 3).setValues(eg);
    escribirFilas_(tvSh, T, [[1, 3], [5, 3]], d.ventasPlatos.map(function (v) {
      return [[new Date(anio, mesIdx_(v.mes), 1), v.nombre, ORIGEN_CLAUDE], [v_(v.unidades), r2v_(v.importe), v_(v.udsLlevar)]];
    }), 'ventas por plato');
    var pares = d.ventasPlatos.filter(function (v) { return v.plato && v.plato !== v.nombre; }).map(function (v) { return [v.nombre, v.plato]; });
    if (pares.length) asignarNombresTpv(pares);
    hecho.push('• ' + d.ventasPlatos.length + ' ventas por plato');
  }
  if (d.nombresTpv && d.nombresTpv.length) {
    asignarNombresTpv(d.nombresTpv);
    hecho.push('• ' + d.nombresTpv.length + ' nombres del TPV relacionados');
  }
  SpreadsheetApp.flush();
  return hecho;
}
var ORIGEN_CLAUDE = 'Claude';   // columna Origen de Ventas TPV para lo cargado con «Cargar datos preparados por Claude»

function v_(x) { return x === undefined || x === null ? '' : x; }
function r2v_(x) { return typeof x === 'number' ? Math.round(x * 100) / 100 : v_(x); }

/** Añade a una lista de Configuración los valores (por su primera columna) que no estén. Devuelve cuántos añadió. */
function anadirALista_(sh, clave, ancho, filas) {
  var d = L.C[clave], c = d.col.charCodeAt(0) - 64, n = d.last - L.C.listFirst + 1;
  var rango = sh.getRange(L.C.listFirst, c, n, ancho), v = rango.getValues(), nuevos = 0;
  filas.forEach(function (f) {
    if (v.some(function (x) { return String(x[0]).toLowerCase() === String(f[0]).toLowerCase(); })) return;
    var libre = v.findIndex(function (x) { return x[0] === ''; });
    if (libre < 0) throw new Error('La lista «' + clave + '» de Configuración está llena.');
    v[libre] = f.concat(['', '', '']).slice(0, ancho); nuevos++;
  });
  if (nuevos) rango.setValues(v);
  return nuevos;
}

/**
 * Escribe filas nuevas en líneas libres de un registro. grupos = [[colInicio, ancho], …] (solo columnas de usuario);
 * cada fila nueva = un array por grupo. Una línea es libre si todos esos grupos están vacíos.
 */
function escribirFilas_(sh, blk, grupos, filas, que) {
  var n = blk.last - blk.first + 1;
  var datos = grupos.map(function (g) { return sh.getRange(blk.first, g[0], n, g[1]).getValues(); });
  var libres = [];
  for (var i = 0; i < n && libres.length < filas.length; i++) {
    if (datos.every(function (m) { return m[i].join('') === ''; })) libres.push(i);
  }
  if (libres.length < filas.length) throw new Error('No caben las ' + filas.length + ' líneas de ' + que + ' (quedan ' + libres.length + ' libres).');
  filas.forEach(function (f, k) { grupos.forEach(function (g, j) { datos[j][libres[k]] = f[j]; }); });
  grupos.forEach(function (g, j) { sh.getRange(blk.first, g[0], n, g[1]).setValues(datos[j]); });
}

/** Como escribirFilas_, pero si ya hay una línea con ese nombre (columna A) la actualiza. */
function actualizarPorNombre_(sh, blk, grupos, filas, que) {
  var n = blk.last - blk.first + 1;
  var datos = grupos.map(function (g) { return sh.getRange(blk.first, g[0], n, g[1]).getValues(); });
  filas.forEach(function (f) {
    var nombre = String(f[0][0]).toLowerCase();
    var i = datos[0].findIndex(function (x) { return String(x[0]).toLowerCase() === nombre; });
    var existe = i >= 0;
    if (!existe) i = datos[0].findIndex(function (x, k) { return datos.every(function (m) { return m[k].join('') === ''; }); });
    if (i < 0) throw new Error('La hoja de ' + que + ' está llena.');
    // En una línea que ya existe, un dato vacío no borra lo que había
    // (y el nombre se queda como estaba escrito)
    grupos.forEach(function (g, j) {
      datos[j][i] = f[j].map(function (v, k) { return existe && (v === '' || (j === 0 && k === 0)) ? datos[j][i][k] : v; });
    });
  });
  grupos.forEach(function (g, j) { sh.getRange(blk.first, g[0], n, g[1]).setValues(datos[j]); });
}
