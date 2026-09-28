/**
 * Estructura de carpetas en el Drive del hostelero, junto a este archivo:
 *   Facturas/<año>/1T (ene-mar) … 4T (oct-dic)
 *   Informes TPV/<año>
 * Se puede ejecutar varias veces: reutiliza lo que ya existe.
 */
var TRIMESTRES = ['1T (ene-mar)', '2T (abr-jun)', '3T (jul-sep)', '4T (oct-dic)'];

function configurarCarpetas() {
  var c = carpetas_(true);
  SpreadsheetApp.getUi().alert('Carpetas listas',
    'En la misma carpeta que este archivo tienes:\n\n• Facturas / ' + c.anio + ' / 1T … 4T\n• Informes TPV / ' + c.anio +
    '\n\nAhí se guardarán las facturas y los informes que envíes desde el panel.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Devuelve (y crea si hace falta) las carpetas del año de la plantilla. */
function carpetas_(crear) {
  var ss = libro_();
  var anio = String(ss.getSheetByName(L.CFG).getRange(L.C.anio).getValue() || new Date().getFullYear());
  var file = DriveApp.getFileById(ss.getId());
  var parents = file.getParents();
  var base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var fact = sub_(base, 'Facturas', crear), factAnio = fact && sub_(fact, anio, crear);
  var tri = TRIMESTRES.map(function (t) { return factAnio && sub_(factAnio, t, crear); });
  var tpv = sub_(base, 'Informes TPV', crear), tpvAnio = tpv && sub_(tpv, anio, crear);
  var props = PropertiesService.getDocumentProperties();
  if (factAnio) props.setProperty('CARPETA_FACTURAS_' + anio, factAnio.getId());
  if (tpvAnio) props.setProperty('CARPETA_TPV_' + anio, tpvAnio.getId());
  return { anio: anio, base: base, facturas: factAnio, trimestres: tri, tpv: tpvAnio };
}

function sub_(parent, name, crear) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return crear ? parent.createFolder(name) : null;
}

/** Carpeta del trimestre para una fecha (la usa el asistente de facturas). */
function carpetaTrimestre_(fecha) {
  var c = carpetas_(true);
  return c.trimestres[Math.floor(fecha.getMonth() / 3)];
}
