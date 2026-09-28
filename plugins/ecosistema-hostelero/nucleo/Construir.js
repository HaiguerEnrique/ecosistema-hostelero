/**
 * Construye la plantilla completa desde cero en la hoja actual.
 * Se ejecuta UNA vez en la plantilla maestra (menú Ecosistema › Administración).
 */

var COLOR = {
  ink: '#1d2b27', brand: '#1f4d43', brandSoft: '#e3eeea', head: '#d5e4de',
  auto: '#f2f4f3', autoInk: '#44504c', line: '#cfd8d4', key: '#eaf3ef',
  warn: '#b3261e', note: '#6b7773'
};
var EUR = '#,##0.00 €;[Red]-#,##0.00 €';
var PCT = '0.0%;[Red]-0.0%';
var FECHA = 'dd/mm/yyyy';

/**
 * La construcción va por etapas porque entera tarda más de los 6 minutos que Apps Script deja por ejecución.
 * Si se acerca el límite, guarda por dónde va (CONSTRUIR_PASO) y al volver a pulsar sigue desde ahí.
 */
var PASOS_CONSTRUIR = [
  function (ss) { construirConfiguracion_(ss.getSheetByName(L.CFG)); }
].concat(L.MESES.map(function (m, i) {
  return function (ss) { construirMes_(ss.getSheetByName(m), i); };
})).concat([
  function (ss) { construirIngredientes_(ss.getSheetByName(L.INGRED)); },
  function (ss) { construirPlatos_(ss.getSheetByName(L.PLATOS)); },
  function (ss) { construirRecetas_(ss.getSheetByName(L.RECETAS)); },
  function (ss) {
    construirVentasTPV_(ss.getSheetByName(L.TPV));
    prepararTpv_(ss); prepararLocales_(ss); hojaDatos_(TPV.PLATOS, TPV.PLATOS_COLS);
  },
  function (ss) { construirResumen_(ss.getSheetByName(L.RES)); },
  function (ss) { construirInicio_(ss.getSheetByName(L.INI)); }
]);
var LIMITE_CONSTRUIR_MS = 4.5 * 60 * 1000;

function construirPlantilla() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getDocumentProperties();
  var inicio = Date.now();
  var paso = props.getProperty('CONSTRUIR_PASO');

  if (paso === null) {
    if (ss.getSheets().length > 1) {
      var r = ui.alert('Reconstruir plantilla',
        'Esto BORRA todas las hojas y datos de este archivo y crea la plantilla desde cero. ¿Seguro?',
        ui.ButtonSet.YES_NO);
      if (r !== ui.Button.YES) return;
    }
    ss.setSpreadsheetTimeZone('Europe/Madrid');
    var orden = [L.INI, L.RES].concat(L.MESES, [L.PLATOS, L.RECETAS, L.INGRED, L.TPV, L.CFG]);
    // Hoja temporal para poder borrar las demás
    var tmp = ss.insertSheet('_tmp_' + Date.now());
    ss.getSheets().forEach(function (s) { if (s.getName() !== tmp.getName()) ss.deleteSheet(s); });
    orden.forEach(function (n, i) { ss.insertSheet(n, i); });
    ss.deleteSheet(tmp);
    props.deleteProperty('DATOS_CLAUDE_CARGADOS');   // hoja nueva: se pueden volver a cargar datos
    paso = 0;
    props.setProperty('CONSTRUIR_PASO', '0');
  }
  paso = Number(paso);

  // Las fórmulas se escriben con sintaxis inglesa (comas); al final se pasa a español y Google las convierte.
  ss.setSpreadsheetLocale('en_US');
  SpreadsheetApp.flush();
  while (paso < PASOS_CONSTRUIR.length) {
    if (Date.now() - inicio > LIMITE_CONSTRUIR_MS) {
      SpreadsheetApp.flush();
      ui.alert('Crear mi plantilla', 'Vamos por la parte ' + paso + ' de ' + PASOS_CONSTRUIR.length +
        '. Google no deja trabajar más de unos minutos seguidos: pulsa otra vez 🍽️ Ecosistema › 🚀 Crear mi plantilla para seguir donde lo dejamos.', ui.ButtonSet.OK);
      return;
    }
    PASOS_CONSTRUIR[paso](ss);
    SpreadsheetApp.flush();
    paso++;
    props.setProperty('CONSTRUIR_PASO', String(paso));
  }

  // Colores de pestañas
  ss.getSheetByName(L.INI).setTabColor(COLOR.brand);
  ss.getSheetByName(L.RES).setTabColor(COLOR.brand);
  L.MESES.forEach(function (m) { ss.getSheetByName(m).setTabColor('#7fa99b'); });
  [L.PLATOS, L.RECETAS, L.INGRED, L.TPV].forEach(function (n) { ss.getSheetByName(n).setTabColor('#c9a227'); });
  ss.getSheetByName(L.CFG).setTabColor('#8a8f8d');

  SpreadsheetApp.flush();
  ss.setSpreadsheetLocale('es_ES');
  props.deleteProperty('CONSTRUIR_PASO');
  props.setProperty('PLANTILLA_VERSION', L.VERSION);
  ss.setActiveSheet(ss.getSheetByName(L.INI));
  SpreadsheetApp.flush();
  ui.alert('Plantilla creada', 'La plantilla v' + L.VERSION + ' está lista. Recarga la página para ver el menú completo de 🍽️ Ecosistema.', ui.ButtonSet.OK);
}

function construccionAMedias_() {
  return PropertiesService.getDocumentProperties().getProperty('CONSTRUIR_PASO') !== null;
}

/* ---------------------------------------------------------------- utilidades */

function titulo_(sh, a1, texto, size) {
  sh.getRange(a1).setValue(texto).setFontSize(size || 18).setFontWeight('bold').setFontColor(COLOR.ink);
}
function seccion_(sh, row, col, ncols, texto) {
  var r = sh.getRange(row, col, 1, ncols);
  if (ncols > 1) r.merge();
  r.setValue(texto).setBackground(COLOR.brand).setFontColor('#ffffff').setFontWeight('bold')
   .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(row, 24);
  return r;
}
function cabecera_(sh, row, col, valores) {
  var r = sh.getRange(row, col, 1, valores.length);
  r.setValues([valores]).setBackground(COLOR.head).setFontWeight('bold').setFontColor(COLOR.ink)
   .setWrap(true).setVerticalAlignment('middle');
  return r;
}
function auto_(range) { return range.setBackground(COLOR.auto).setFontColor(COLOR.autoInk); }
function bordes_(range) {
  return range.setBorder(true, true, true, true, true, true, COLOR.line, SpreadsheetApp.BorderStyle.SOLID);
}
/** Rellena una columna con la misma fórmula, sustituyendo {r} por el nº de fila. */
function formulaCol_(sh, col, first, last, plantilla) {
  var arr = [];
  for (var r = first; r <= last; r++) arr.push([plantilla.replace(/\{r\}/g, r)]);
  sh.getRange(col + first + ':' + col + last).setFormulas(arr);
}
function lista_(rangoFuente, ayuda) {
  return SpreadsheetApp.newDataValidation().requireValueInRange(rangoFuente, true)
    .setAllowInvalid(false).setHelpText(ayuda).build();
}
function proteger_(range, desc) {
  range.protect().setDescription(desc || 'Celda automática').setWarningOnly(true);
}
function ss_() { return SpreadsheetApp.getActive(); }

/* ----------------------------------------------------------- Configuración */

function construirConfiguracion_(sh) {
  sh.clear();
  sh.setHiddenGridlines(true);
  titulo_(sh, 'A1', 'Configuración');
  sh.getRange('A2').setValue('Rellena los datos de tu negocio. Las listas de abajo alimentan los desplegables de toda la plantilla: puedes añadir o renombrar, pero no dejes huecos en medio.')
    .setFontColor(COLOR.note).setFontStyle('italic');

  seccion_(sh, 3, 1, 3, 'DATOS DEL NEGOCIO');
  var datos = [
    ['Nombre del local', ''],
    ['Año de esta plantilla', new Date().getFullYear()],
    ['Saldo inicial del año (€)', ''],
    ['Inventario inicial del año (€)', ''],
    ['IVA de ventas por defecto', 0.10]
  ];
  sh.getRange('A4:B8').setValues(datos);
  sh.getRange('A4:A8').setFontWeight('bold');
  bordes_(sh.getRange('A4:B8'));
  sh.getRange('B6:B7').setNumberFormat(EUR);
  sh.getRange('B8').setNumberFormat('0%');
  sh.getRange('B5').setNumberFormat('0');
  sh.getRange('C6').setValue('Dinero total en caja + bancos el 1 de enero.').setFontColor(COLOR.note);
  sh.getRange('C7').setValue('Déjalo vacío si no haces inventario.').setFontColor(COLOR.note);

  seccion_(sh, 10, 1, 3, 'REFERENCIAS DE INDICADORES (orientativas: ajústalas a tu negocio)');
  cabecera_(sh, 11, 1, ['Indicador', 'Objetivo', 'Alerta a partir de']);
  sh.getRange('A12:C16').setValues([
    ['Food cost (coste de mercancía / ventas)', 0.30, 0.35],
    ['Personal / ventas', 0.35, 0.40],
    ['Prime cost (mercancía + personal)', 0.60, 0.65],
    ['Margen operativo (mínimo deseado)', 0.10, 0],
    ['Food cost por plato', 0.30, 0.35]
  ]);
  sh.getRange('B12:C16').setNumberFormat('0%');
  bordes_(sh.getRange('A11:C16'));
  sh.getRange('D15').setValue('En margen: por debajo del objetivo = vigilar; por debajo de la alerta = pérdidas.').setFontColor(COLOR.note);

  var f = L.C.listFirst;
  seccion_(sh, 19, 1, 3, 'CATEGORÍAS DE GASTO');
  cabecera_(sh, 20, 1, ['Categoría', 'Tipo', 'Qué incluye']);
  var cats = [
    ['Compras de mercancía (sin desglosar)', 'Mercancía', 'Facturas mixtas: comida, bebida, envases… en la misma factura (Makro, cash & carry).'],
    ['Compras comida', 'Mercancía', 'Proveedores solo de alimentación.'],
    ['Compras bebida', 'Mercancía', 'Distribuidoras de bebida, cerveza, vino, refrescos.'],
    ['Envases y desechables', 'Mercancía', 'Cajas, bolsas y envases de delivery y para llevar.'],
    ['Nóminas', 'Personal', 'Sueldo neto pagado a trabajadores.'],
    ['Seguridad Social (empresa)', 'Personal', 'Seguros sociales (TC1).'],
    ['Cuota de autónomo', 'Personal', 'Cuota mensual de autónomos del titular o socios.'],
    ['Otros gastos de personal', 'Personal', 'Uniformes, formación, extras, dietas.'],
    ['Alquiler', 'Explotación', 'Alquiler del local y comunidad.'],
    ['Suministros (luz, agua, gas)', 'Explotación', ''],
    ['Teléfono e internet', 'Explotación', ''],
    ['Gestoría y asesoría', 'Explotación', ''],
    ['Seguros', 'Explotación', 'Seguro del local, responsabilidad civil.'],
    ['Software y TPV', 'Explotación', 'TPV, programas, suscripciones.'],
    ['Marketing y publicidad', 'Explotación', 'Anuncios, diseño, redes, cartelería.'],
    ['Mantenimiento y reparaciones', 'Explotación', ''],
    ['Limpieza e higiene', 'Explotación', 'Productos de limpieza, control de plagas.'],
    ['Menaje y utensilios', 'Explotación', 'Vajilla, cubertería, utensilios de cocina.'],
    ['Comisiones bancarias y datáfono', 'Explotación', ''],
    ['Comisiones de plataformas delivery', 'Explotación', 'Glovo, Uber Eats, Just Eat: comisión que se quedan.'],
    ['Transporte y combustible', 'Explotación', ''],
    ['Tasas y tributos locales', 'Explotación', 'Basuras, terraza, IAE, licencias.'],
    ['Otros gastos', 'Explotación', ''],
    ['Maquinaria y equipamiento', 'Inversión', 'Compras que duran años: maquinaria, mobiliario, reformas.'],
    ['Préstamos y leasing (cuotas)', 'Financiación', 'Cuota mensual de préstamos, renting y leasing.'],
    ['Impuestos (IVA, IRPF, Sociedades)', 'Impuestos', 'Pagos a Hacienda: modelos 303, 111, 115, 130, 200…'],
    ['Retirada del propietario', 'Propietario', 'Dinero que se lleva el dueño. No es gasto del negocio.']
  ];
  sh.getRange(f, 1, cats.length, 3).setValues(cats);
  bordes_(sh.getRange('A20:C' + L.C.cat.last));

  seccion_(sh, 19, 5, 2, 'TIPOS DE GASTO (no modificar)');
  cabecera_(sh, 20, 5, ['Tipo', 'Qué significa']);
  var tipos = [
    ['Mercancía', 'Lo que vendes: comida, bebida y envases. Forma el coste de mercancía (CMV).'],
    ['Personal', 'Nóminas, Seguridad Social y cuota de autónomo.'],
    ['Explotación', 'Gastos para funcionar: alquiler, luz, gestoría, marketing…'],
    ['Inversión', 'Compras que duran años. No restan del resultado del mes, sí de la caja.'],
    ['Financiación', 'Cuotas de préstamos y leasing. Salen de caja, no son gasto operativo.'],
    ['Impuestos', 'IVA, IRPF, Sociedades. El IVA no es gasto: ya está descontado de las ventas.'],
    ['Propietario', 'Lo que te llevas tú. No es gasto del negocio.']
  ];
  sh.getRange(f, 5, tipos.length, 2).setValues(tipos);
  bordes_(sh.getRange('E20:F' + L.C.tipos.last));
  sh.getRange('E' + f + ':E' + L.C.tipos.last).setFontWeight('bold');
  proteger_(sh.getRange('E20:F' + L.C.tipos.last), 'Tipos de gasto: la plantilla depende de estos nombres');

  seccion_(sh, 19, 8, 2, 'CANALES DE VENTA');
  cabecera_(sh, 20, 8, ['Canal', 'Tipo de canal']);
  var canales = [['Sala', 'Directo'], ['Terraza', 'Directo'], ['Para llevar', 'Directo'],
                 ['Reparto propio', 'Directo'], ['Glovo', 'Plataforma'], ['Uber Eats', 'Plataforma'],
                 ['Just Eat', 'Plataforma'], ['Catering y eventos', 'Directo'], ['Otros', 'Directo']];
  sh.getRange(f, 8, canales.length, 2).setValues(canales);
  bordes_(sh.getRange('H20:I' + L.C.canal.last));
  sh.getRange('I' + f + ':I' + L.C.canal.last).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Directo', 'Plataforma'], true).setAllowInvalid(false).build());
  sh.getRange('H' + (L.C.canal.last + 1)).setValue('Plataforma = lo cobra la plataforma (no entra en tu caja ese día).').setFontColor(COLOR.note);

  seccion_(sh, 19, 11, 2, 'CATEGORÍAS DE INGRESO');
  cabecera_(sh, 20, 11, ['Categoría', 'IVA']);
  sh.getRange(f, 11, 3, 2).setValues([['Ventas del día', 0.10], ['Eventos y catering', 0.10], ['Otros ingresos', 0.21]]);
  sh.getRange('L' + f + ':L' + L.C.catIng.last).setNumberFormat('0%');
  bordes_(sh.getRange('K20:L' + L.C.catIng.last));

  seccion_(sh, 19, 14, 1, 'FORMAS DE PAGO');
  cabecera_(sh, 20, 14, ['Forma de pago']);
  sh.getRange(f, 14, 6, 1).setValues([['Tarjeta'], ['Transferencia'], ['Domiciliación'], ['Efectivo'], ['Bizum'], ['Pendiente']]);
  bordes_(sh.getRange('N20:N' + L.C.pago.last));
  sh.getRange('N' + (L.C.pago.last + 1)).setValue('"Pendiente" = factura aún sin pagar.').setFontColor(COLOR.note);

  seccion_(sh, 19, 16, 1, 'CATEGORÍAS DE LA CARTA');
  cabecera_(sh, 20, 16, ['Categoría']);
  sh.getRange(f, 16, 6, 1).setValues([['Entrantes'], ['Principales'], ['Postres'], ['Bebidas'], ['Menús'], ['Otros']]);
  bordes_(sh.getRange('P20:P' + L.C.carta.last));

  sh.setColumnWidth(1, 290); sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 330); sh.setColumnWidth(4, 20);
  sh.setColumnWidth(5, 110); sh.setColumnWidth(6, 380); sh.setColumnWidth(7, 20);
  sh.setColumnWidth(8, 150); sh.setColumnWidth(9, 110); sh.setColumnWidth(10, 20);
  sh.setColumnWidth(11, 160); sh.setColumnWidth(12, 70); sh.setColumnWidth(13, 20);
  sh.setColumnWidth(14, 140); sh.setColumnWidth(15, 20); sh.setColumnWidth(16, 150);
  sh.getRange('B' + f + ':B' + L.C.cat.last).setDataValidation(
    lista_(sh.getRange('E' + f + ':E' + L.C.tipos.last), 'Elige un tipo de la lista'));
  sh.setFrozenRows(0);
}

/* ------------------------------------------------------------------- Meses */

function construirMes_(sh, i) {
  var mes = L.MESES[i], m = i + 1, M = L.M, I = L.ING, G = L.GAS;
  var prev = i > 0 ? SH(L.MESES[i - 1]) : null;
  sh.clear();
  sh.getRange('A:Z').clearDataValidations();

  var widths = [44, 88, 120, 150, 170, 170, 96, 104, 104, 104, 112, 112, 96, 96, 200];
  widths.forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });

  sh.getRange('A1').setFormula('="' + mes + ' "&' + CF(L.C.anio) + '&IF(' + CF(L.C.nombre) + '="",""," · "&' + CF(L.C.nombre) + ')')
    .setFontSize(18).setFontWeight('bold').setFontColor(COLOR.ink);
  sh.getRange('A2').setValue('Casillas blancas: las rellenas tú.  Casillas grises: se calculan solas.  En "Día" escribe solo el número (ej. 5).')
    .setFontColor(COLOR.note).setFontStyle('italic');

  /* ---- Bloque A: cuenta de resultados (etiqueta A:C, € en D, % en E) ---- */
  seccion_(sh, 4, 1, 5, 'CUENTA DE RESULTADOS DEL MES');
  cabecera_(sh, 5, 1, ['Concepto', '', '', '€', '% s/ventas']);
  sh.getRange('A5:C5').merge();
  var tipoSum = function (t) { return '=SUMIF(' + col_('G', G) + ',"' + t + '",' + col_('H', G) + ')'; };
  var pl = [
    [6, 'Ventas sin IVA', '=SUM(' + col_('K', I) + ')', 'key'],
    [7, 'Compras de mercancía', tipoSum('Mercancía')],
    [8, 'Variación de inventario', '=IF(' + M.invFin + '="",0,' + M.invIni + '-' + M.invFin + ')'],
    [9, 'Coste de mercancía (CMV)', '=D7+D8'],
    [10, 'MARGEN BRUTO', '=D6-D9', 'key'],
    [11, 'Personal', tipoSum('Personal')],
    [12, 'Gastos de explotación', tipoSum('Explotación')],
    [13, 'RESULTADO OPERATIVO', '=D10-D11-D12', 'key'],
    [16, 'Inversiones', tipoSum('Inversión')],
    [17, 'Préstamos y financiación', tipoSum('Financiación')],
    [18, 'Impuestos (IVA, IRPF, Sociedades)', tipoSum('Impuestos')],
    [19, 'Retirada del propietario', tipoSum('Propietario')],
    [21, 'CAJA GENERADA EN EL MES', '=' + M.cobrado + '-' + M.pagado, 'key']
  ];
  pl.forEach(function (p) {
    var r = p[0];
    sh.getRange(r, 1, 1, 3).merge().setValue(p[1]);
    sh.getRange('D' + r).setFormula(p[2]).setNumberFormat(EUR);
    sh.getRange('E' + r).setFormula('=IF($D$6=0,"",D' + r + '/$D$6)').setNumberFormat(PCT);
    auto_(sh.getRange('A' + r + ':E' + r));
    if (p[3] === 'key') sh.getRange('A' + r + ':E' + r).setFontWeight('bold').setBackground(COLOR.key).setFontColor(COLOR.ink);
  });
  sh.getRange('E6').setValue('');
  sh.getRange('A15:E15').merge().setValue('Pagos que NO son gasto del negocio (salen de caja pero no restan del resultado)')
    .setFontColor(COLOR.note).setFontStyle('italic').setFontSize(9);
  sh.getRange('A22:E22').merge().setValue('Caja generada = total cobrado − total pagado (con IVA, incluye inversiones, préstamos, impuestos y retiradas).')
    .setFontColor(COLOR.note).setFontStyle('italic').setFontSize(9).setWrap(true);
  bordes_(sh.getRange('A5:E13')); bordes_(sh.getRange('A16:E19')); bordes_(sh.getRange('A21:E21'));

  seccion_(sh, 23, 1, 5, 'INDICADORES CLAVE');
  cabecera_(sh, 24, 1, ['Indicador', '', '', 'Valor', 'Estado']);
  sh.getRange('A24:C24').merge();
  var est = function (v, o, a, malo) {
    return '=IF(' + v + '="","—",IF(' + v + '<=' + CF(o) + ',"✅ Bien",IF(' + v + '<=' + CF(a) + ',"⚠️ Vigilar","🔴 ' + malo + '")))';
  };
  var kpis = [
    [25, 'Food cost (CMV / ventas)', '=IF(D6=0,"",D9/D6)', est('D25', L.C.fcObj, L.C.fcAlerta, 'Alto')],
    [26, 'Personal / ventas', '=IF(D6=0,"",D11/D6)', est('D26', L.C.perObj, L.C.perAlerta, 'Alto')],
    [27, 'Prime cost (CMV + personal) / ventas', '=IF(D6=0,"",(D9+D11)/D6)', est('D27', L.C.primeObj, L.C.primeAlerta, 'Alto')],
    [28, 'Margen operativo', '=IF(D6=0,"",D13/D6)',
      '=IF(D28="","—",IF(D28>=' + CF(L.C.margenObj) + ',"✅ Bien",IF(D28>=' + CF(L.C.margenAlerta) + ',"⚠️ Vigilar","🔴 Pérdidas")))'],
    [29, 'Explotación / ventas', '=IF(D6=0,"",D12/D6)', '="—"']
  ];
  kpis.forEach(function (k) {
    sh.getRange(k[0], 1, 1, 3).merge().setValue(k[1]);
    sh.getRange('D' + k[0]).setFormula(k[2]).setNumberFormat(PCT);
    sh.getRange('E' + k[0]).setFormula(k[3]);
  });
  auto_(sh.getRange('A25:E29')); bordes_(sh.getRange('A24:E29'));
  sh.getRange('A30:E30').merge().setValue('Referencias editables en Configuración. Son orientativas.')
    .setFontColor(COLOR.note).setFontStyle('italic').setFontSize(9);

  /* ---- Bloque B: gastos por categoría (F categoría, G tipo, H €, I %) ---- */
  seccion_(sh, 4, 6, 4, 'GASTOS POR CATEGORÍA');
  cabecera_(sh, 5, 6, ['Categoría', 'Tipo', '€ sin IVA', '% s/ventas']);
  var fB = [], cfgRow = L.C.listFirst;
  for (var r = M.catFirst; r <= M.catLast; r++, cfgRow++) {
    fB.push([
      '=IF(' + SH(L.CFG) + '$A' + cfgRow + '="","",' + SH(L.CFG) + '$A' + cfgRow + ')',
      '=IF(F' + r + '="","",' + SH(L.CFG) + '$B' + cfgRow + ')',
      '=IF(F' + r + '="","",SUMIF(' + col_('F', G) + ',F' + r + ',' + col_('H', G) + '))',
      '=IF(OR(F' + r + '="",$D$6=0),"",H' + r + '/$D$6)'
    ]);
  }
  sh.getRange('F' + M.catFirst + ':I' + M.catLast).setFormulas(fB);
  sh.getRange('F34').setValue('Total');
  sh.getRange('H34').setFormula('=SUM(H' + M.catFirst + ':H' + M.catLast + ')');
  sh.getRange('I34').setFormula('=IF($D$6=0,"",H34/$D$6)');
  sh.getRange('F35').setValue('Sin categoría válida');
  sh.getRange('H35').setFormula('=SUM(' + col_('H', G) + ')-H34');
  sh.getRange('H' + M.catFirst + ':H35').setNumberFormat(EUR);
  sh.getRange('I' + M.catFirst + ':I34').setNumberFormat(PCT);
  auto_(sh.getRange('F' + M.catFirst + ':I35')); bordes_(sh.getRange('F5:I35'));
  sh.getRange('F34:I34').setFontWeight('bold').setBackground(COLOR.key);
  sh.getRange('F' + M.catFirst + ':F35').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  /* ---- Bloque C: caja e inventario (J:K etiqueta, L valor) ---- */
  seccion_(sh, 4, 10, 3, 'CAJA');
  cabecera_(sh, 5, 10, ['Concepto', '', '€']); sh.getRange('J5:K5').merge();
  var saldoIni = prev ? '=IF(' + prev + M.saldoReal + '<>"",' + prev + M.saldoReal + ',' + prev + M.saldoPrev + ')' : '=' + CF(L.C.saldoIni);
  var invIni = prev ? '=IF(' + prev + M.invFin + '<>"",' + prev + M.invFin + ',' + prev + M.invIni + ')' : '=' + CF(L.C.invIni);
  var caja = [
    [6, 'Saldo inicial del mes', saldoIni],
    [7, '+ Total cobrado', '=SUM(' + col_('F', I) + ')'],
    [8, '− Total pagado', '=SUMIF(' + col_('K', G) + ',"<>Pendiente",' + col_('J', G) + ')'],
    [9, 'Saldo previsto a fin de mes', '=L6+L7-L8'],
    [10, 'Efectivo en caja (real)', null],
    [11, 'Bancos (real)', null],
    [12, 'Saldo real', '=IF(COUNT(L10:L11)=0,"",SUM(L10:L11))'],
    [13, 'Diferencia (real − previsto)', '=IF(L12="","",L12-L9)'],
    [14, 'Pendiente de pago a proveedores', '=SUMIF(' + col_('K', G) + ',"Pendiente",' + col_('J', G) + ')'],
    [17, 'Efectivo', '=SUM(' + col_('I', I) + ')'],
    [18, 'Tarjeta', '=SUM(' + col_('G', I) + ')'],
    [19, 'Bizum / transferencia', '=SUM(' + col_('H', I) + ')'],
    [20, 'Plataformas delivery', '=SUM(' + col_('J', I) + ')'],
    [23, 'Inventario inicial', invIni],
    [24, 'Inventario final (lo cuentas tú)', null]
  ];
  caja.forEach(function (c) {
    sh.getRange(c[0], 10, 1, 2).merge().setValue(c[1]);
    var v = sh.getRange('L' + c[0]).setNumberFormat(EUR);
    if (c[2]) { v.setFormula(c[2]); auto_(sh.getRange('J' + c[0] + ':L' + c[0])); }
    else { sh.getRange('J' + c[0] + ':K' + c[0]).setBackground(COLOR.auto); v.setBackground('#ffffff').setFontWeight('bold'); }
  });
  sh.getRange('J9:L9').setFontWeight('bold').setBackground(COLOR.key);
  sh.getRange('J12:L13').setFontWeight('bold');
  seccion_(sh, 16, 10, 3, 'COBROS POR FORMA DE PAGO');
  seccion_(sh, 22, 10, 3, 'INVENTARIO');
  sh.getRange('J25:L25').merge().setValue('Si no haces inventario deja el final vacío: el CMV será igual a las compras.')
    .setFontColor(COLOR.note).setFontStyle('italic').setFontSize(9).setWrap(true);
  sh.setRowHeight(25, 30);
  bordes_(sh.getRange('J5:L14')); bordes_(sh.getRange('J17:L20')); bordes_(sh.getRange('J23:L24'));

  /* ---- Bloque D: datos operativos y punto de equilibrio (M:N etiqueta, O valor) ---- */
  seccion_(sh, 4, 13, 3, 'DATOS OPERATIVOS');
  cabecera_(sh, 5, 13, ['Concepto', '', 'Valor']); sh.getRange('M5:N5').merge();
  var dia = col_('A', I);
  var ope = [
    [6, 'Días con ventas', formulaDias_(dia, col_('F', I)), '0'],
    [7, 'Tickets', '=IF(SUM(' + col_('M', I) + ')=0,"",SUM(' + col_('M', I) + '))', '#,##0'],
    [8, 'Comensales', '=IF(SUM(' + col_('N', I) + ')=0,"",SUM(' + col_('N', I) + '))', '#,##0'],
    [9, 'Ticket medio (IVA incl.)', '=IF(O7="","",L7/O7)', EUR],
    [10, 'Gasto por comensal (IVA incl.)', '=IF(O8="","",L7/O8)', EUR],
    [11, 'Venta media por día (sin IVA)', '=IF(OR(O6="",O6=0),"",D6/O6)', EUR],
    [14, 'Costes fijos (personal + explotación)', '=D11+D12', EUR],
    [15, 'Margen de contribución', '=IF(D6=0,"",1-D9/D6)', PCT],
    [16, 'Venta necesaria al mes (sin IVA)', '=IF(OR(O15="",O15<=0),"",O14/O15)', EUR],
    [17, 'Venta necesaria por día', '=IF(OR(O16="",O6="",O6=0),"",O16/O6)', EUR],
    [18, 'Ventas reales − venta necesaria', '=IF(O16="","",D6-O16)', EUR]
  ];
  ope.forEach(function (o) {
    sh.getRange(o[0], 13, 1, 2).merge().setValue(o[1]);
    sh.getRange('O' + o[0]).setFormula(o[2]).setNumberFormat(o[3]);
    auto_(sh.getRange('M' + o[0] + ':O' + o[0]));
  });
  seccion_(sh, 13, 13, 3, 'PUNTO DE EQUILIBRIO');
  sh.getRange('M19:O20').merge().setValue('Cuánto necesitas vender para no perder dinero con tus costes fijos y tu margen actual. Tickets y comensales son opcionales: si tu TPV no los da, quedan vacíos.')
    .setFontColor(COLOR.note).setFontStyle('italic').setFontSize(9).setWrap(true).setVerticalAlignment('top');
  bordes_(sh.getRange('M5:O11')); bordes_(sh.getRange('M14:O18'));
  sh.getRange('M16:O16').setFontWeight('bold').setBackground(COLOR.key);

  /* ---- Registro de ingresos ---- */
  seccion_(sh, I.title, 1, 15, 'REGISTRO DE INGRESOS  ·  una línea por día y canal (o por cierre de caja)');
  cabecera_(sh, I.hdr, 1, L.ING_COLS);
  sh.setRowHeight(I.hdr, 36);
  var fecha = function (r) {
    return '=IF(A' + r + '="","",IF(DAY(DATE(' + CF(L.C.anio) + ',' + m + ',A' + r + '))<>A' + r +
      ',"Día no válido",DATE(' + CF(L.C.anio) + ',' + m + ',A' + r + ')))';
  };
  var tipoCanal = 'IFERROR(VLOOKUP(D{r},' + cfgList_('canal', 2) + ',2,FALSE),"Directo")';
  formulaCol_(sh, 'B', I.first, I.last, fecha('{r}'));
  formulaCol_(sh, 'I', I.first, I.last, '=IF(F{r}="","",IF(' + tipoCanal + '="Plataforma",0,F{r}-N(G{r})-N(H{r})))');
  formulaCol_(sh, 'J', I.first, I.last, '=IF(F{r}="","",IF(' + tipoCanal + '="Plataforma",F{r},0))');
  formulaCol_(sh, 'K', I.first, I.last, '=IF(F{r}="","",ROUND(F{r}/(1+IFERROR(VLOOKUP(E{r},' + cfgList_('catIng', 2) + ',2,FALSE),' + CF(L.C.ivaVentas) + ')),2))');
  formulaCol_(sh, 'L', I.first, I.last, '=IF(F{r}="","",F{r}-K{r})');
  var ingAll = sh.getRange(I.first, 1, I.last - I.first + 1, 15);
  bordes_(sh.getRange(I.hdr, 1, I.last - I.hdr + 1, 15));
  ['B', 'I', 'J', 'K', 'L'].forEach(function (c) { auto_(sh.getRange(c + I.first + ':' + c + I.last)); });
  sh.getRange('B' + I.first + ':B' + I.last).setNumberFormat(FECHA);
  sh.getRange('F' + I.first + ':L' + I.last).setNumberFormat(EUR);
  sh.getRange('M' + I.first + ':N' + I.last).setNumberFormat('0');

  /* ---- Registro de gastos ---- */
  seccion_(sh, G.title, 1, 13, 'REGISTRO DE GASTOS Y FACTURAS  ·  300 líneas');
  cabecera_(sh, G.hdr, 1, L.GAS_COLS);
  sh.setRowHeight(G.hdr, 36);
  formulaCol_(sh, 'B', G.first, G.last, fecha('{r}'));
  formulaCol_(sh, 'G', G.first, G.last, '=IF(F{r}="","",IFERROR(VLOOKUP(F{r},' + cfgList_('cat', 2) + ',2,FALSE),"Revisar"))');
  formulaCol_(sh, 'J', G.first, G.last, '=IF(AND(H{r}="",I{r}=""),"",N(H{r})+N(I{r}))');
  bordes_(sh.getRange(G.hdr, 1, G.last - G.hdr + 1, 13));
  ['B', 'G', 'J'].forEach(function (c) { auto_(sh.getRange(c + G.first + ':' + c + G.last)); });
  sh.getRange('B' + G.first + ':B' + G.last).setNumberFormat(FECHA);
  sh.getRange('H' + G.first + ':J' + G.last).setNumberFormat(EUR);

  /* ---- Validaciones ---- */
  var cfg = ss_().getSheetByName(L.CFG), lf = L.C.listFirst;
  var diaVal = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 31).setAllowInvalid(false)
    .setHelpText('Escribe solo el día del mes (1 a 31). La fecha se completa sola.').build();
  sh.getRange('A' + I.first + ':A' + I.last).setDataValidation(diaVal);
  sh.getRange('A' + G.first + ':A' + G.last).setDataValidation(diaVal);
  sh.getRange('D' + I.first + ':D' + I.last).setDataValidation(lista_(cfg.getRange('H' + lf + ':H' + L.C.canal.last), 'Canal de venta. Se editan en Configuración.'));
  sh.getRange('E' + I.first + ':E' + I.last).setDataValidation(lista_(cfg.getRange('K' + lf + ':K' + L.C.catIng.last), 'Categoría de ingreso. Se editan en Configuración.'));
  sh.getRange('F' + G.first + ':F' + G.last).setDataValidation(lista_(cfg.getRange('A' + lf + ':A' + L.C.cat.last), 'Categoría de gasto. Si falta una, añádela en Configuración.'));
  sh.getRange('K' + G.first + ':K' + G.last).setDataValidation(lista_(cfg.getRange('N' + lf + ':N' + L.C.pago.last), 'Forma de pago. "Pendiente" si aún no la has pagado.'));
  var num = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Importe en euros').build();
  sh.getRange('F' + I.first + ':H' + I.last).setDataValidation(num);
  sh.getRange('H' + G.first + ':I' + G.last).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberBetween(-1000000, 1000000).setAllowInvalid(false)
      .setHelpText('Importe en euros. Usa negativo para abonos.').build());

  /* ---- Formato condicional ---- */
  var red = SpreadsheetApp.newConditionalFormatRule;
  var rules = [
    red().whenFormulaSatisfied('=AND($L$13<>"",ABS($L$13)>=1)').setFontColor(COLOR.warn).setBackground('#fdecea').setRanges([sh.getRange('J13:L13')]).build(),
    red().whenTextEqualTo('Día no válido').setFontColor(COLOR.warn).setBold(true).setRanges([sh.getRange('B' + I.first + ':B' + I.last), sh.getRange('B' + G.first + ':B' + G.last)]).build(),
    red().whenTextEqualTo('Revisar').setFontColor(COLOR.warn).setBold(true).setRanges([sh.getRange('G' + G.first + ':G' + G.last)]).build(),
    red().whenNumberLessThan(0).setFontColor(COLOR.warn).setBackground('#fdecea').setRanges([sh.getRange('I' + I.first + ':I' + I.last)]).build(),
    red().whenFormulaSatisfied('=$H$35<>0').setFontColor(COLOR.warn).setBold(true).setRanges([sh.getRange('F35:I35')]).build(),
    red().whenTextEqualTo('Pendiente').setFontColor('#9a6700').setBackground('#fff4d6').setRanges([sh.getRange('K' + G.first + ':K' + G.last)]).build()
  ];
  sh.setConditionalFormatRules(rules);

  /* ---- Protección suave (avisa antes de sobrescribir fórmulas) ---- */
  proteger_(sh.getRange('A1:I35'), 'Resumen del mes (automático)');
  proteger_(sh.getRange('J4:L9'), 'Caja (automático)');
  proteger_(sh.getRange('J12:L23'), 'Caja e inventario (automático)');
  proteger_(sh.getRange('M4:O20'), 'Datos operativos (automático)');
  proteger_(sh.getRange('B' + I.first + ':B' + I.last), 'Fecha automática');
  proteger_(sh.getRange('I' + I.first + ':L' + I.last), 'Cálculos de ingresos');
  proteger_(sh.getRange('B' + G.first + ':B' + G.last), 'Fecha automática');
  proteger_(sh.getRange('G' + G.first + ':G' + G.last), 'Tipo automático');
  proteger_(sh.getRange('J' + G.first + ':J' + G.last), 'Total automático');

  sh.getRange('A4:O35').setVerticalAlignment('middle');
  sh.getRange('M6:N18').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange('J6:K24').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.setFrozenRows(0);
}

/* ------------------------------------------------------------ Escandallos */

function construirIngredientes_(sh) {
  var R = L.IN;
  sh.clear();
  titulo_(sh, 'A1', 'Ingredientes');
  sh.getRange('A2').setValue('Un ingrediente por línea, con el precio al que lo compras. El coste por unidad se calcula solo. La merma es lo que se pierde al limpiar o cocinar (ej. 20%).')
    .setFontColor(COLOR.note).setFontStyle('italic');
  cabecera_(sh, 4, 1, ['Ingrediente', 'Unidad', 'Proveedor', 'Precio del formato (sin IVA)', 'Cantidad del formato (en la unidad)',
    'Coste por unidad', 'Merma %', 'Coste real por unidad', 'Actualizado', 'Notas']);
  sh.setRowHeight(4, 40);
  formulaCol_(sh, 'F', R.first, R.last, '=IF(OR(D{r}="",E{r}="",N(E{r})=0),"",D{r}/E{r})');
  formulaCol_(sh, 'H', R.first, R.last, '=IF(F{r}="","",IF(N(G{r})>=1,"Merma no válida",F{r}/(1-N(G{r}))))');
  ['F', 'H'].forEach(function (c) { auto_(sh.getRange(c + R.first + ':' + c + R.last)); });
  sh.getRange('B' + R.first + ':B' + R.last).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['kg', 'l', 'ud'], true).setAllowInvalid(false).setHelpText('kg, l o ud').build());
  sh.getRange('D' + R.first + ':D' + R.last).setNumberFormat(EUR);
  sh.getRange('E' + R.first + ':E' + R.last).setNumberFormat('#,##0.000');
  sh.getRange('F' + R.first + ':F' + R.last).setNumberFormat('#,##0.0000 €');
  sh.getRange('H' + R.first + ':H' + R.last).setNumberFormat('#,##0.0000 €');
  sh.getRange('G' + R.first + ':G' + R.last).setNumberFormat('0%');
  sh.getRange('I' + R.first + ':I' + R.last).setNumberFormat(FECHA);
  bordes_(sh.getRange(4, 1, R.last - 3, 10));
  [220, 70, 150, 120, 130, 110, 80, 120, 100, 240].forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });
  sh.setFrozenRows(4);
  proteger_(sh.getRange('F' + R.first + ':F' + R.last)); proteger_(sh.getRange('H' + R.first + ':H' + R.last));
}

function construirRecetas_(sh) {
  var R = L.RE, P = L.PL, N = L.IN;
  sh.clear();
  titulo_(sh, 'A1', 'Recetas (escandallos)');
  sh.getRange('A2').setValue('Una línea por ingrediente de cada plato. La cantidad va en la misma unidad que el ingrediente (kg, l o ud). Ej.: 0,150 kg de carne.')
    .setFontColor(COLOR.note).setFontStyle('italic');
  cabecera_(sh, 4, 1, ['Plato', 'Ingrediente', 'Cantidad', 'Unidad', 'Coste por unidad', 'Coste de la línea']);
  formulaCol_(sh, 'D', R.first, R.last, '=IF(B{r}="","",IFERROR(VLOOKUP(B{r},' + SH(L.INGRED) + '$A$' + N.first + ':$B$' + N.last + ',2,FALSE),"?"))');
  formulaCol_(sh, 'E', R.first, R.last, '=IF(B{r}="","",IFERROR(VLOOKUP(B{r},' + SH(L.INGRED) + '$A$' + N.first + ':$H$' + N.last + ',8,FALSE),"Falta ingrediente"))');
  formulaCol_(sh, 'F', R.first, R.last, '=IF(OR(C{r}="",NOT(ISNUMBER(E{r}))),"",C{r}*E{r})');
  ['D', 'E', 'F'].forEach(function (c) { auto_(sh.getRange(c + R.first + ':' + c + R.last)); });
  var pl = ss_().getSheetByName(L.PLATOS), ing = ss_().getSheetByName(L.INGRED);
  sh.getRange('A' + R.first + ':A' + R.last).setDataValidation(lista_(pl.getRange('A' + P.first + ':A' + P.last), 'Plato dado de alta en la hoja Platos'));
  sh.getRange('B' + R.first + ':B' + R.last).setDataValidation(lista_(ing.getRange('A' + N.first + ':A' + N.last), 'Ingrediente dado de alta en la hoja Ingredientes'));
  sh.getRange('C' + R.first + ':C' + R.last).setNumberFormat('#,##0.000');
  sh.getRange('E' + R.first + ':E' + R.last).setNumberFormat('#,##0.0000 €');
  sh.getRange('F' + R.first + ':F' + R.last).setNumberFormat('#,##0.000 €');
  bordes_(sh.getRange(4, 1, R.last - 3, 6));
  [220, 220, 90, 70, 120, 120].forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });
  sh.setFrozenRows(4);
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Falta ingrediente')
    .setFontColor(COLOR.warn).setRanges([sh.getRange('E' + R.first + ':E' + R.last)]).build()]);
  proteger_(sh.getRange('D' + R.first + ':F' + R.last));
}

function construirPlatos_(sh) {
  var P = L.PL, R = L.RE, T = L.TV;
  sh.clear();
  titulo_(sh, 'A1', 'Platos y rentabilidad de la carta');
  sh.getRange('A2').setValue('Da de alta cada plato con su precio. El coste sale de Recetas y las unidades vendidas de Ventas TPV. La clasificación compara cada plato con los de su misma categoría.')
    .setFontColor(COLOR.note).setFontStyle('italic');
  cabecera_(sh, 4, 1, ['Plato', 'Categoría', 'PVP (IVA incl.)', 'IVA (vacío = por defecto)', 'Envase delivery (€)',
    'Coste de la ración', 'PVP sin IVA', 'Food cost %', 'Margen por plato', 'Food cost delivery %',
    'Estado', 'Unidades vendidas', 'Margen total', 'Clasificación']);
  sh.setRowHeight(4, 40);
  var rng = function (c) { return '$' + c + '$' + P.first + ':$' + c + '$' + P.last; };
  var catU = 'SUMIF(' + rng('B') + ',B{r},' + rng('L') + ')';
  formulaCol_(sh, 'F', P.first, P.last, '=IF(A{r}="","",SUMIF(' + SH(L.RECETAS) + '$A$' + R.first + ':$A$' + R.last + ',A{r},' + SH(L.RECETAS) + '$F$' + R.first + ':$F$' + R.last + '))');
  formulaCol_(sh, 'G', P.first, P.last, '=IF(OR(A{r}="",C{r}=""),"",C{r}/(1+IF(D{r}="",' + CF(L.C.ivaVentas) + ',D{r})))');
  formulaCol_(sh, 'H', P.first, P.last, '=IF(OR(G{r}="",N(G{r})=0),"",F{r}/G{r})');
  formulaCol_(sh, 'I', P.first, P.last, '=IF(G{r}="","",G{r}-F{r})');
  formulaCol_(sh, 'J', P.first, P.last, '=IF(OR(G{r}="",N(G{r})=0),"",(F{r}+N(E{r}))/G{r})');
  formulaCol_(sh, 'K', P.first, P.last, '=IF(A{r}="","",IF(N(F{r})=0,"Falta receta",IF(H{r}="","Falta precio",IF(H{r}<=' + CF(L.C.fcPlatoObj) + ',"✅ Bien",IF(H{r}<=' + CF(L.C.fcPlatoAlerta) + ',"⚠️ Vigilar","🔴 Alto")))))');
  formulaCol_(sh, 'L', P.first, P.last, '=IF(A{r}="","",SUMIF(' + SH(L.TPV) + '$D$' + T.first + ':$D$' + T.last + ',A{r},' + SH(L.TPV) + '$E$' + T.first + ':$E$' + T.last + '))');
  formulaCol_(sh, 'M', P.first, P.last, '=IF(OR(A{r}="",I{r}=""),"",I{r}*N(L{r}))');
  formulaCol_(sh, 'N', P.first, P.last,
    '=IF(A{r}="","",IF(OR(N(L{r})=0,I{r}="",B{r}=""),"Sin datos",' +
    'IF(L{r}/' + catU + '>=0.7/COUNTIFS(' + rng('B') + ',B{r},' + rng('L') + ',">0"),' +
    'IF(I{r}>=SUMIF(' + rng('B') + ',B{r},' + rng('M') + ')/' + catU + ',"⭐ Estrella","🐴 Caballo"),' +
    'IF(I{r}>=SUMIF(' + rng('B') + ',B{r},' + rng('M') + ')/' + catU + ',"🧩 Enigma","🐶 Perro"))))');
  ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'].forEach(function (c) { auto_(sh.getRange(c + P.first + ':' + c + P.last)); });
  var cfg = ss_().getSheetByName(L.CFG);
  sh.getRange('B' + P.first + ':B' + P.last).setDataValidation(lista_(cfg.getRange('P' + L.C.listFirst + ':P' + L.C.carta.last), 'Categoría de la carta (Configuración)'));
  sh.getRange('C' + P.first + ':C' + P.last).setNumberFormat(EUR);
  sh.getRange('D' + P.first + ':D' + P.last).setNumberFormat('0%');
  sh.getRange('E' + P.first + ':G' + P.last).setNumberFormat(EUR);
  sh.getRange('H' + P.first + ':H' + P.last).setNumberFormat(PCT);
  sh.getRange('I' + P.first + ':I' + P.last).setNumberFormat(EUR);
  sh.getRange('J' + P.first + ':J' + P.last).setNumberFormat(PCT);
  sh.getRange('L' + P.first + ':L' + P.last).setNumberFormat('#,##0');
  sh.getRange('M' + P.first + ':M' + P.last).setNumberFormat(EUR);
  bordes_(sh.getRange(4, 1, P.last - 3, 14));
  [220, 120, 100, 90, 90, 100, 100, 90, 100, 100, 110, 90, 110, 120].forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });
  sh.setFrozenRows(4); sh.setFrozenColumns(1);
  proteger_(sh.getRange('F' + P.first + ':N' + P.last));
  sh.getRange('P4').setValue('Cómo leer la clasificación').setFontWeight('bold');
  sh.getRange('P5:P8').setValues([['⭐ Estrella: se vende mucho y deja buen margen. Cuídalo.'],
    ['🐴 Caballo: se vende mucho pero deja poco. Sube precio o baja coste.'],
    ['🧩 Enigma: deja buen margen pero se vende poco. Destácalo en la carta.'],
    ['🐶 Perro: se vende poco y deja poco. Replantéalo o quítalo.']]).setFontColor(COLOR.note);
  sh.setColumnWidth(16, 420);
}

function construirVentasTPV_(sh) {
  var T = L.TV, MP = L.MAP, P = L.PL;
  sh.clear();
  titulo_(sh, 'A1', 'Ventas del TPV por plato');
  sh.getRange('A2').setValue('Aquí se vuelcan los informes de ventas por producto de tu TPV (el panel lo hace automáticamente). Si el nombre del TPV no coincide con el del plato, relaciónalos en la tabla de la derecha.')
    .setFontColor(COLOR.note).setFontStyle('italic');
  cabecera_(sh, 4, 1, ['Mes (fecha)', 'Nombre en el TPV', 'Canal', 'Plato de la carta', 'Unidades', 'Importe (IVA incl.)']);
  formulaCol_(sh, 'D', T.first, T.last, '=IF(B{r}="","",IFERROR(VLOOKUP(B{r},$I$' + MP.first + ':$J$' + MP.last + ',2,FALSE),IF(COUNTIF(' + SH(L.PLATOS) + '$A$' + P.first + ':$A$' + P.last + ',B{r})>0,B{r},"Sin asignar")))');
  auto_(sh.getRange('D' + T.first + ':D' + T.last));
  sh.getRange('A' + T.first + ':A' + T.last).setNumberFormat('mm/yyyy');
  sh.getRange('E' + T.first + ':E' + T.last).setNumberFormat('#,##0');
  sh.getRange('F' + T.first + ':F' + T.last).setNumberFormat(EUR);
  bordes_(sh.getRange(4, 1, T.last - 3, 6));
  cabecera_(sh, 4, 9, ['Nombre en el TPV', 'Plato de la carta']);
  sh.getRange('J' + MP.first + ':J' + MP.last).setDataValidation(lista_(ss_().getSheetByName(L.PLATOS).getRange('A' + P.first + ':A' + P.last), 'Plato de la hoja Platos'));
  bordes_(sh.getRange(4, 9, MP.last - 3, 2));
  [90, 220, 110, 220, 80, 110, 20, 20, 220, 220].forEach(function (w, k) { sh.setColumnWidth(k + 1, w); });
  sh.setFrozenRows(4);
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Sin asignar')
    .setFontColor(COLOR.warn).setRanges([sh.getRange('D' + T.first + ':D' + T.last)]).build()]);
  proteger_(sh.getRange('D' + T.first + ':D' + T.last));
}

/* ----------------------------------------------------------- Resumen anual */

function construirResumen_(sh) {
  var M = L.M;
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.getRange('A1').setFormula('="Resumen anual "&' + CF(L.C.anio) + '&IF(' + CF(L.C.nombre) + '="",""," · "&' + CF(L.C.nombre) + ')')
    .setFontSize(18).setFontWeight('bold').setFontColor(COLOR.ink);
  sh.getRange('A2').setValue('Todo se calcula solo a partir de las hojas de cada mes. Aquí no se escribe nada.').setFontColor(COLOR.note).setFontStyle('italic');
  var mesesCorto = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  var row = 4;

  // Crea una tabla mes a mes. filas: [etiqueta, celda del mes | función(mes)->fórmula, total('sum'|'last'|'first'|fórmula|null), formato, clave?]
  function tabla(tituloTxt, filas, conPct) {
    seccion_(sh, row, 1, conPct ? 15 : 14, tituloTxt); row++;
    cabecera_(sh, row, 1, ['Concepto'].concat(mesesCorto, ['Total año'], conPct ? ['% s/ventas'] : []));
    var hdrRow = row; row++;
    var start = row;
    filas.forEach(function (f) {
      var fs = L.MESES.map(function (mes) {
        return typeof f[1] === 'function' ? f[1](mes) : '=' + SH(mes) + f[1];
      });
      sh.getRange(row, 1).setValue(f[0]);
      if (typeof f[0] === 'string' && f[0].charAt(0) === '=') sh.getRange(row, 1).setFormula(f[0]);
      sh.getRange(row, 2, 1, 12).setFormulas([fs]);
      var tot = f[2], tf = '';
      if (tot === 'sum') tf = '=SUM(B' + row + ':M' + row + ')';
      else if (tot === 'first') tf = '=B' + row;
      else if (tot === 'last') tf = '=M' + row;
      else if (tot) tf = tot.replace(/\{r\}/g, row);
      if (tf) sh.getRange(row, 14).setFormula(tf);
      if (conPct) sh.getRange(row, 15).setFormula('=IF(OR(N' + row + '="",$N$' + (start) + '=0),"",N' + row + '/$N$' + start + ')').setNumberFormat(PCT);
      sh.getRange(row, 2, 1, 13).setNumberFormat(f[3] || EUR);
      if (f[4]) sh.getRange(row, 1, 1, conPct ? 15 : 14).setFontWeight('bold').setBackground(COLOR.key);
      row++;
    });
    bordes_(sh.getRange(hdrRow, 1, row - hdrRow, conPct ? 15 : 14));
    row++;
    return start;
  }

  var plStart = tabla('CUENTA DE RESULTADOS', [
    ['Ventas sin IVA', M.ventas, 'sum', EUR, true],
    ['Compras de mercancía', M.compras, 'sum'],
    ['Variación de inventario', M.varInv, 'sum'],
    ['Coste de mercancía (CMV)', M.cmv, 'sum'],
    ['MARGEN BRUTO', M.margenBruto, 'sum', EUR, true],
    ['Personal', M.personal, 'sum'],
    ['Gastos de explotación', M.explotacion, 'sum'],
    ['RESULTADO OPERATIVO', M.resultado, 'sum', EUR, true],
    ['Inversiones', M.inversiones, 'sum'],
    ['Préstamos y financiación', M.financiacion, 'sum'],
    ['Impuestos (IVA, IRPF, Sociedades)', M.impuestos, 'sum'],
    ['Retirada del propietario', M.retirada, 'sum'],
    ['CAJA GENERADA', M.cajaGenerada, 'sum', EUR, true]
  ], true);
  sh.getRange(plStart, 15).setValue('');
  var v = 'N' + plStart; // ventas totales
  tabla('INDICADORES', [
    ['Food cost (CMV / ventas)', M.kFoodCost, '=IF(' + v + '=0,"",N' + (plStart + 3) + '/' + v + ')', PCT],
    ['Personal / ventas', M.kPersonal, '=IF(' + v + '=0,"",N' + (plStart + 5) + '/' + v + ')', PCT],
    ['Prime cost', M.kPrime, '=IF(' + v + '=0,"",(N' + (plStart + 3) + '+N' + (plStart + 5) + ')/' + v + ')', PCT, true],
    ['Margen operativo', M.kMargen, '=IF(' + v + '=0,"",N' + (plStart + 7) + '/' + v + ')', PCT, true],
    ['Explotación / ventas', M.kExplot, '=IF(' + v + '=0,"",N' + (plStart + 6) + '/' + v + ')', PCT]
  ]);
  var cajaStart = tabla('CAJA E INVENTARIO', [
    ['Saldo inicial', M.saldoIni, 'first'],
    ['Total cobrado', M.cobrado, 'sum'],
    ['Total pagado', M.pagado, 'sum'],
    ['Saldo previsto a fin de mes', M.saldoPrev, 'last', EUR, true],
    ['Saldo real contado', M.saldoReal, null],
    ['Diferencia', M.diferencia, 'sum'],
    ['Pendiente de pago', M.pendiente, null],
    ['Inventario final', M.invFin, null]
  ]);
  tabla('COBROS POR FORMA DE PAGO', [
    ['Efectivo', M.cobEfectivo, 'sum'], ['Tarjeta', M.cobTarjeta, 'sum'],
    ['Bizum / transferencia', M.cobBizum, 'sum'], ['Plataformas delivery', M.cobPlataformas, 'sum']
  ]);
  var opStart = tabla('DATOS OPERATIVOS', [
    ['Días con ventas', M.dias, 'sum', '0'],
    ['Tickets', M.tickets, 'sum', '#,##0'],
    ['Comensales', M.comensales, 'sum', '#,##0'],
    ['Ticket medio (IVA incl.)', M.ticketMedio, null],
    ['Venta media por día (sin IVA)', M.ventaDia, null],
    ['Venta necesaria (punto de equilibrio)', M.ventaNecesaria, 'sum', EUR, true]
  ]);
  // Totales anuales que no son una simple suma
  sh.getRange(opStart + 3, 14).setFormula('=IF(N' + (opStart + 1) + '=0,"",N' + (cajaStart + 1) + '/N' + (opStart + 1) + ')');
  sh.getRange(opStart + 4, 14).setFormula('=IF(N' + opStart + '=0,"",' + v + '/N' + opStart + ')');

  // Gastos por categoría
  var catRows = [];
  for (var r = M.catFirst; r <= M.catLast; r++) {
    (function (rr) {
      catRows.push(['=' + SH(L.MESES[0]) + 'F' + rr, function (mes) { return '=' + SH(mes) + 'H' + rr; }, '=IF($A{r}="","",SUM(B{r}:M{r}))']);
    })(r);
  }
  var catStart = tabla('GASTOS POR CATEGORÍA (sin IVA)', catRows);
  for (var k = 0; k < catRows.length; k++) {
    var rr = catStart + k;
    sh.getRange(rr, 15).setFormula('=IF(OR($A' + rr + '="",' + v + '=0),"",N' + rr + '/' + v + ')').setNumberFormat(PCT);
  }
  sh.getRange(catStart - 1, 15).setValue('% s/ventas').setBackground(COLOR.head).setFontWeight('bold');

  // Ventas por canal y por categoría de ingreso
  var I = L.ING;
  var canalRows = [];
  for (var c = L.C.listFirst; c <= L.C.canal.last; c++) {
    (function (cc, idx) {
      canalRows.push(['=' + CF('H' + cc).replace(/\$/g, ''), function (mes) {
        return '=IF($A' + '{row}' + '="","",SUMIF(' + SH(mes) + col_('D', I) + ',$A{row},' + SH(mes) + col_('K', I) + '))';
      }, '=IF($A{r}="","",SUM(B{r}:M{r}))']);
    })(c);
  }
  var canalStart = row + 2;
  canalRows.forEach(function (f, idx) {
    var fn = f[1];
    f[1] = function (mes) { return fn(mes).replace(/\{row\}/g, canalStart + idx); };
  });
  tabla('VENTAS POR CANAL (sin IVA)', canalRows);

  var catIngRows = [];
  for (var c2 = L.C.listFirst; c2 <= L.C.catIng.last; c2++) {
    catIngRows.push(['=' + CF('K' + c2).replace(/\$/g, ''), null, '=IF($A{r}="","",SUM(B{r}:M{r}))']);
  }
  var ciStart = row + 2;
  catIngRows.forEach(function (f, idx) {
    f[1] = function (mes) {
      return '=IF($A' + (ciStart + idx) + '="","",SUMIF(' + SH(mes) + col_('E', I) + ',$A' + (ciStart + idx) + ',' + SH(mes) + col_('K', I) + '))';
    };
  });
  tabla('VENTAS POR CATEGORÍA DE INGRESO (sin IVA)', catIngRows);

  // IVA
  var G = L.GAS;
  var ivaStart = tabla('IVA (orientativo: confírmalo con tu gestoría)', [
    ['IVA repercutido (ventas)', function (mes) { return '=SUM(' + SH(mes) + col_('L', I) + ')'; }, 'sum'],
    ['IVA soportado (gastos)', function (mes) { return '=SUM(' + SH(mes) + col_('I', G) + ')'; }, 'sum'],
    ['Diferencia', function (mes) { return '=0'; }, 'sum', EUR, true]
  ]);
  var difRow = ivaStart + 2;
  var colL = 'BCDEFGHIJKLM';
  for (var mi = 0; mi < 12; mi++) {
    var cl = colL.charAt(mi);
    sh.getRange(cl + difRow).setFormula('=' + cl + ivaStart + '-' + cl + (ivaStart + 1));
  }
  seccion_(sh, row, 1, 5, 'IVA POR TRIMESTRE (modelo 303 orientativo)'); row++;
  cabecera_(sh, row, 1, ['Trimestre', 'IVA repercutido', 'IVA soportado', 'Resultado', 'A pagar / a compensar']); var tq = row; row++;
  var tri = [['1T (ene-mar)', 'B', 'D'], ['2T (abr-jun)', 'E', 'G'], ['3T (jul-sep)', 'H', 'J'], ['4T (oct-dic)', 'K', 'M']];
  tri.forEach(function (t) {
    sh.getRange(row, 1).setValue(t[0]);
    sh.getRange(row, 2).setFormula('=SUM(' + t[1] + ivaStart + ':' + t[2] + ivaStart + ')');
    sh.getRange(row, 3).setFormula('=SUM(' + t[1] + (ivaStart + 1) + ':' + t[2] + (ivaStart + 1) + ')');
    sh.getRange(row, 4).setFormula('=B' + row + '-C' + row);
    sh.getRange(row, 5).setFormula('=IF(D' + row + '>=0,"A pagar","A compensar")');
    sh.getRange(row, 2, 1, 3).setNumberFormat(EUR);
    row++;
  });
  bordes_(sh.getRange(tq, 1, 5, 5));
  row++;

  // Datos para el gráfico (ocultos a la derecha)
  sh.getRange('R4:U4').setValues([['Mes', 'Ventas sin IVA', 'Mercancía + personal', 'Resultado operativo']]);
  var chartRows = [];
  for (var j = 0; j < 12; j++) {
    var cc2 = colL.charAt(j);
    chartRows.push([mesesCorto[j], '=' + cc2 + plStart, '=' + cc2 + (plStart + 3) + '+' + cc2 + (plStart + 5), '=' + cc2 + (plStart + 7)]);
  }
  sh.getRange('R5:R16').setValues(chartRows.map(function (x) { return [x[0]]; }));
  sh.getRange('S5:U16').setFormulas(chartRows.map(function (x) { return [x[1], x[2], x[3]]; }));
  sh.getRange('R4:U16').setFontColor('#999999').setFontSize(8);
  crearGrafico_(sh);

  sh.setColumnWidth(1, 270);
  for (var w = 2; w <= 13; w++) sh.setColumnWidth(w, 92);
  sh.setColumnWidth(14, 110); sh.setColumnWidth(15, 90); sh.setColumnWidth(16, 24);
  proteger_(sh.getRange('A1:U' + row), 'Resumen anual (automático)');
}

/* ------------------------------------------------------------------ Inicio */

function construirInicio_(sh) {
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 28); sh.setColumnWidth(2, 760);
  sh.getRange('B2').setFormula('="Control financiero "&' + CF(L.C.anio) + '&IF(' + CF(L.C.nombre) + '="",""," · "&' + CF(L.C.nombre) + ')')
    .setFontSize(22).setFontWeight('bold').setFontColor(COLOR.ink);
  sh.getRange('B3').setValue('Plantilla del Ecosistema Hostelero · v' + L.VERSION).setFontColor(COLOR.note);
  var bloques = [
    ['CÓMO EMPEZAR', [
      '1.  Ve a la pestaña Configuración y rellena el nombre del local, el año, el saldo inicial y (si lo tienes) el inventario inicial.',
      '2.  Revisa las listas de Configuración: categorías de gasto, canales, formas de pago. Adáptalas a tu negocio.',
      '3.  Cada mes tiene su pestaña. Apunta los ingresos (una línea por día y canal) y los gastos (una línea por factura).',
      '4.  En la columna "Día" escribe solo el número del día: la fecha completa se rellena sola.',
      '5.  A final de mes cuenta el efectivo y el saldo de bancos (y el inventario si lo haces) y apúntalo en el bloque CAJA.',
      '6.  El mes siguiente arranca solo con el saldo y el inventario del mes anterior.'
    ]],
    ['COLORES', [
      'Casillas blancas: las rellenas tú.',
      'Casillas grises: se calculan solas. Si intentas escribir en ellas, la hoja te avisará.',
      'Rojo: algo que revisar (día que no existe, categoría no válida, descuadre de caja).'
    ]],
    ['LAS PESTAÑAS', [
      'Resumen anual: la cuenta de resultados del año mes a mes, indicadores, caja, IVA por trimestre y gráfico.',
      'Enero … Diciembre: el día a día. Arriba el resumen del mes; debajo, el registro de ingresos y el de gastos.',
      'Platos, Recetas e Ingredientes: tus escandallos. Food cost y margen de cada plato.',
      'Ventas TPV: las ventas por plato de tu TPV. Con ellas se clasifica la carta (estrellas, caballos, enigmas y perros).',
      'Configuración: los datos de tu negocio y las listas de la plantilla.'
    ]],
    ['CÓMO SE LEE EL RESULTADO', [
      'MARGEN BRUTO = ventas − coste de la mercancía.',
      'RESULTADO OPERATIVO = margen bruto − personal − gastos de explotación. Es lo que gana de verdad el negocio.',
      'Inversiones, préstamos, impuestos y tu retirada NO restan del resultado: salen de caja, pero no son gasto del negocio.',
      'CAJA GENERADA = todo lo cobrado − todo lo pagado. Es lo que realmente entra o sale del banco.'
    ]],
    ['COMPRAS MEZCLADAS (tipo Makro)', [
      'Si en una misma factura hay comida, bebida y otros, usa "Compras de mercancía (sin desglosar)". El food cost sale igual de bien.',
      'Si quieres el detalle, divide la factura en varias líneas (una por categoría) con el mismo nº de factura. El asistente del panel lo hará por ti.'
    ]]
  ];
  var r = 5;
  bloques.forEach(function (b) {
    sh.getRange(r, 2).setValue(b[0]).setBackground(COLOR.brand).setFontColor('#ffffff').setFontWeight('bold');
    sh.setRowHeight(r, 26); r++;
    b[1].forEach(function (t) { sh.getRange(r, 2).setValue(t).setWrap(true).setFontColor(COLOR.ink); r++; });
    r++;
  });
  sh.getRange(r, 2).setValue('Los cálculos de IVA son orientativos: confirma siempre con tu gestoría.').setFontColor(COLOR.note).setFontStyle('italic');
}

/** Gráfico del Resumen anual a partir de la tabla auxiliar R4:U16. */
function crearGrafico_(sh) {
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  var chart = sh.newChart().setChartType(Charts.ChartType.COMBO)
    .addRange(sh.getRange('R4:U16'))
    .setNumHeaders(1)
    .setOption('title', 'Ventas, costes principales y resultado por mes')
    .setOption('series', { 0: { type: 'bars', color: '#7fa99b' }, 1: { type: 'bars', color: '#c9a227' }, 2: { type: 'line', color: '#1f4d43', lineWidth: 3, pointSize: 5 } })
    .setOption('legend', { position: 'bottom' })
    .setOption('vAxis', { format: '#,##0 €' })
    .setOption('width', 760).setOption('height', 340)
    .setPosition(4, 17, 0, 0).build();
  sh.insertChart(chart);
}

/** Rehace solo el gráfico (Administración). */
function rehacerGrafico() { crearGrafico_(SpreadsheetApp.getActive().getSheetByName(L.RES)); }
