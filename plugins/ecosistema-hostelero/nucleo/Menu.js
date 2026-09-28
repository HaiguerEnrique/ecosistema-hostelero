/** Menú "Ecosistema" y atajos de navegación. */

function onOpen() {
  try { aplicarParches_(); } catch (e) { console.warn('Parches: ' + e.message); }
  var ui = SpreadsheetApp.getUi();
  var aMedias = construccionAMedias_();
  if (!SpreadsheetApp.getActive().getSheetByName(L.CFG) || aMedias) {   // hoja recién creada o a medio construir
    ui.createMenu('🍽️ Ecosistema').addItem(aMedias ? '🚀 Crear mi plantilla (seguir)' : '🚀 Crear mi plantilla', 'construirPlantilla').addToUi();
    return;
  }
  var menu = ui.createMenu('🍽️ Ecosistema')
    .addItem('📊 Abrir panel y asistente', 'abrirPanel')
    .addItem('📱 Panel en el móvil', 'enlaceMovil')
    .addSeparator()
    .addItem('📅 Ir al mes actual', 'irMesActual')
    .addItem('➡️ Ir a ingresos de este mes', 'irIngresos')
    .addItem('➡️ Ir a gastos de este mes', 'irGastos')
    .addSeparator()
    .addItem('🔑 Configurar clave de IA', 'configurarClave')
    .addItem('📁 Configurar carpetas de Drive', 'configurarCarpetas')
    .addItem('🗓️ Crear archivo del año siguiente', 'crearAnioSiguiente')
    .addItem('📥 Cargar datos preparados por Claude', 'cargarDatosClaude')
    .addSeparator()
    .addSubMenu(ui.createMenu('Administración')
      .addItem('Construir plantilla desde cero (borra todo)', 'construirPlantilla')
      .addItem('Rehacer gráfico del resumen', 'rehacerGrafico')
      .addItem('Actualizar plantilla', 'actualizarPlantilla'));
  // Personalizaciones del hostelero (archivos Personal*.js, que las actualizaciones no tocan)
  if (typeof menuPersonal_ === 'function') {
    try { menuPersonal_(ui, menu); } catch (e) { console.warn('Menú personal: ' + e.message); }
  }
  menu.addToUi();
}

function hojaMes_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  if (L.MESES.indexOf(sh.getName()) >= 0) return sh;
  return ss.getSheetByName(L.MESES[new Date().getMonth()]);
}
function irMesActual() {
  var sh = SpreadsheetApp.getActive().getSheetByName(L.MESES[new Date().getMonth()]);
  sh.activate(); sh.setActiveRange(sh.getRange('A1'));
}
function irIngresos() { irPrimeraLibre_(L.ING); }
function irGastos() { irPrimeraLibre_(L.GAS); }
function irPrimeraLibre_(blk) {
  var sh = hojaMes_();
  sh.activate();
  var vals = sh.getRange('A' + blk.first + ':A' + blk.last).getValues();
  var r = blk.first;
  for (var i = vals.length - 1; i >= 0; i--) { if (vals[i][0] !== '') { r = blk.first + i + 1; break; } }
  r = Math.min(r, blk.last);
  sh.setActiveRange(sh.getRange('A' + blk.hdr));   // lleva la cabecera a la vista
  sh.setActiveRange(sh.getRange('A' + r));
}
