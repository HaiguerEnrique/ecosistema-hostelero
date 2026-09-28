/**
 * Actualizaciones de la plantilla. Cada parche corrige hojas ya creadas (fórmulas, formatos)
 * sin tocar los datos del usuario. Se aplican solos al abrir el archivo, en orden y una sola vez.
 * Al añadir un parche, sube también L.VERSION y refleja el cambio en Construir.js.
 */
var PARCHES = [
  { version: '2.0.1', descripcion: 'Días con ventas vacío si hay ventas sin día apuntado', formulas: true,
    aplicar: function (ss) {
      var I = L.ING, A = col_('A', I), F = col_('F', I);
      L.MESES.forEach(function (m) {
        var sh = ss.getSheetByName(m);
        sh.getRange(L.M.dias).setFormula(formulaDias_(A, F));
        sh.getRange(L.M.ventaDia).setFormula('=IF(OR(O6="",O6=0),"",D6/O6)');
        sh.getRange(L.M.ventaNecDia).setFormula('=IF(OR(O16="",O6="",O6=0),"",O16/O6)');
      });
    } },
  { version: '2.1.0', descripcion: 'Importador del TPV: unidades para llevar, coste de bolsa por pedido y pestañas de datos del TPV',
    aplicar: function (ss) { prepararTpv_(ss); } },
  { version: '2.2.0', descripcion: 'Varios locales en un panel: tabla «Locales del grupo» en Configuración',
    aplicar: function (ss) { prepararLocales_(ss); } },
  { version: '2.3.0', descripcion: 'Informes del TPV por día (se puede subir el cierre de cada noche) e ingresos automáticos desde el TPV',
    aplicar: function (ss) { prepararTpvDiario_(ss); } },
  { version: '2.5.0', descripcion: 'Configuración: «Datos del TPV desde» (lo anterior a esa fecha se ignora al subir informes)',
    aplicar: function (ss) { prepararTpvDesde_(ss); } }
];

function formulaDias_(A, F) {
  return '=IF(COUNTIFS(' + F + ',"<>",' + A + ',"")>0,"",SUMPRODUCT((' + A + '<>"")/COUNTIF(' + A + ',' + A + '&"")))';
}

function versionMenor_(a, b) {
  var x = a.split('.').map(Number), y = b.split('.').map(Number);
  for (var i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) < (y[i] || 0); }
  return false;
}

/** Aplica los parches pendientes. Devuelve la lista de los aplicados. */
function aplicarParches_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(L.CFG) || construccionAMedias_()) return [];   // sin plantilla o a medio construir
  var props = PropertiesService.getDocumentProperties();
  var actual = props.getProperty('PLANTILLA_VERSION') || '2.0.0';
  var aplicados = [];
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return [];
  try {
    PARCHES.forEach(function (p) {
      if (!versionMenor_(actual, p.version)) return;
      // Las fórmulas se escriben con sintaxis inglesa: se cambia el idioma durante el parche.
      var locale = ss.getSpreadsheetLocale(), cambiar = p.formulas && locale !== 'en_US';
      if (cambiar) ss.setSpreadsheetLocale('en_US');
      try { p.aplicar(ss); SpreadsheetApp.flush(); }
      finally { if (cambiar) ss.setSpreadsheetLocale(locale); }
      props.setProperty('PLANTILLA_VERSION', p.version);
      actual = p.version;
      aplicados.push(p.version + ' · ' + p.descripcion);
    });
  } finally { lock.releaseLock(); }
  return aplicados;
}

/** Menú › Administración: aplica manualmente y enseña el resultado. */
function actualizarPlantilla() {
  var hechos = aplicarParches_();
  SpreadsheetApp.getUi().alert('Actualizar plantilla',
    hechos.length ? 'Aplicado:\n\n' + hechos.join('\n') : 'La plantilla ya está al día (v' +
      (PropertiesService.getDocumentProperties().getProperty('PLANTILLA_VERSION') || '2.0.0') + ').',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Estructura para los informes del TPV (la usan el constructor y el parche 2.1.0). */
function prepararTpv_(ss) {
  var tv = ss.getSheetByName(L.TPV);
  tv.getRange('C4').setValue('Origen');
  tv.getRange('G4').setValue('Uds. para llevar').setBackground(COLOR.head).setFontWeight('bold').setWrap(true);
  tv.setColumnWidth(7, 90);
  tv.getRange('G' + L.TV.first + ':G' + L.TV.last).setNumberFormat('#,##0');
  var cfg = ss.getSheetByName(L.CFG);
  cfg.getRange('A9').setValue('Coste de bolsa por pedido para llevar (€)').setFontWeight('bold');
  if (cfg.getRange('B9').getValue() === '') cfg.getRange('B9').setValue(0);
  cfg.getRange('B9').setNumberFormat(EUR);
  cfg.getRange('C9').setValue('Bolsa o envase que se da por pedido de llevar/domicilio (no por plato). Se suma al CMV teórico.').setFontColor(COLOR.note);
  bordes_(cfg.getRange('A4:B9'));
  [[TPV.DIAS, TPV.DIAS_COLS], [TPV.HORAS, TPV.HORAS_COLS], [TPV.CAM, TPV.CAM_COLS]].forEach(function (h) { hojaDatos_(h[0], h[1]); });
}

/** 2.5.0: casilla «Datos del TPV desde» en Configuración (fila 17, libre bajo los indicadores). */
function prepararTpvDesde_(ss) {
  var cfg = ss.getSheetByName(L.CFG), a = L.C.tpvDesde.replace('B', 'A'), c = L.C.tpvDesde.replace('B', 'C');
  cfg.getRange(a).setValue('Datos del TPV desde (fecha)').setFontWeight('bold');
  cfg.getRange(L.C.tpvDesde).setNumberFormat(FECHA);
  cfg.getRange(c).setValue('Opcional. Al subir informes del TPV se ignoran las noches anteriores a esta fecha (por ejemplo, si tus cuentas empiezan más tarde que el TPV).')
    .setFontColor(COLOR.note).setWrap(true);
  bordes_(cfg.getRange(a + ':' + L.C.tpvDesde));
}

/** 2.3.0: las hojas de horas y camareros pasan a ser por día; nueva hoja de platos por día. */
function prepararTpvDiario_(ss) {
  [[TPV.HORAS, TPV.HORAS_COLS], [TPV.CAM, TPV.CAM_COLS]].forEach(function (h) {
    var sh = ss.getSheetByName(h[0]);
    if (sh && sh.getRange('A1').getValue() === 'Mes') {
      sh.clearContents();
      sh.getRange(1, 1, 1, h[1].length).setValues([h[1]]).setFontWeight('bold').setBackground('#d5e4de').setWrap(true);
      sh.getRange('A2:A').setNumberFormat('dd/mm/yyyy');
    }
  });
  hojaDatos_(TPV.PLATOS, TPV.PLATOS_COLS);
}
