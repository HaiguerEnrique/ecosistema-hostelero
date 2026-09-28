/**
 * Crea el archivo del año siguiente: copia esta plantilla, la vacía de movimientos,
 * y arrastra el saldo de caja y el inventario de diciembre.
 * Mantiene configuración, platos, recetas, ingredientes y el mapeo del TPV.
 */
function crearAnioSiguiente() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var cfg = ss.getSheetByName(L.CFG);
  var anio = Number(cfg.getRange(L.C.anio).getValue()) || new Date().getFullYear();
  var nuevo = anio + 1;
  var r = ui.alert('Crear ' + nuevo,
    'Se creará una copia de este archivo para ' + nuevo + ' en la misma carpeta, vacía de movimientos, ' +
    'con el saldo de caja y el inventario de diciembre como punto de partida. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  var dic = ss.getSheetByName('Diciembre');
  var saldo = dic.getRange(L.M.saldoReal).getValue();
  if (saldo === '') saldo = dic.getRange(L.M.saldoPrev).getValue();
  var inv = dic.getRange(L.M.invFin).getValue();
  if (inv === '') inv = dic.getRange(L.M.invIni).getValue();

  var file = DriveApp.getFileById(ss.getId());
  var parents = file.getParents();
  var carpeta = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var nombre = file.getName().indexOf(String(anio)) >= 0 ? file.getName().replace(String(anio), String(nuevo)) : file.getName() + ' ' + nuevo;
  var copia = file.makeCopy(nombre, carpeta);
  var dst = SpreadsheetApp.openById(copia.getId());

  dst.getSheetByName(L.CFG).getRange(L.C.anio).setValue(nuevo);
  dst.getSheetByName(L.CFG).getRange(L.C.saldoIni).setValue(saldo);
  dst.getSheetByName(L.CFG).getRange(L.C.invIni).setValue(inv === '' ? '' : inv);
  vaciarMovimientos_(dst);

  ui.alert('Listo', 'Nuevo archivo creado: ' + nombre + '\n\n' + copia.getUrl() +
    '\n\nÁbrelo y usa "Configurar carpetas de Drive" para crear las carpetas de ' + nuevo + '.' +
    ' Si tienes varios locales, cambia en Configuración › Locales del grupo los enlaces por los de las hojas de ' + nuevo + ' de los demás.', ui.ButtonSet.OK);
}

/** Borra solo lo que escribe el usuario; fórmulas, validaciones y formato se mantienen. */
function vaciarMovimientos_(ss) {
  var I = L.ING, G = L.GAS;
  L.MESES.forEach(function (m) {
    var sh = ss.getSheetByName(m);
    sh.getRangeList([
      'A' + I.first + ':A' + I.last, 'C' + I.first + ':H' + I.last, 'M' + I.first + ':O' + I.last,
      'A' + G.first + ':A' + G.last, 'C' + G.first + ':F' + G.last, 'H' + G.first + ':I' + G.last, 'K' + G.first + ':M' + G.last,
      L.M.efectivoReal, L.M.bancosReal, L.M.invFin
    ]).clearContent();
  });
  var T = L.TV;
  ss.getSheetByName(L.TPV).getRangeList(['A' + T.first + ':C' + T.last, 'E' + T.first + ':G' + T.last]).clearContent();
  // Datos del TPV por día del año anterior
  [TPV.DIAS, TPV.HORAS, TPV.CAM, TPV.PLATOS].forEach(function (nombre) {
    var sh = ss.getSheetByName(nombre);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  });
}
