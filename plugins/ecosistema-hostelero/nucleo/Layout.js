/**
 * Ecosistema Hostelero — mapa de la plantilla.
 * Todas las piezas (plantilla, panel, chat) leen las posiciones desde aquí.
 * Si cambias una fila o columna, cámbiala SOLO aquí.
 */
var L = {
  VERSION: '2.6.1',

  INI: 'Inicio',
  RES: 'Resumen anual',
  CFG: 'Configuración',
  PLATOS: 'Platos',
  RECETAS: 'Recetas',
  INGRED: 'Ingredientes',
  TPV: 'Ventas TPV',
  MESES: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio',
          'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],

  // Registros en cada hoja de mes
  ING: { title: 37, hdr: 38, first: 39, last: 188 },   // 150 líneas de ingresos
  GAS: { title: 191, hdr: 192, first: 193, last: 492 }, // 300 líneas de gastos

  // Columnas del registro de ingresos
  ING_COLS: ['Día', 'Fecha', 'Concepto', 'Canal', 'Categoría', 'Ventas brutas (IVA incl.)',
             'Tarjeta', 'Bizum / transferencia', 'Efectivo', 'Plataformas',
             'Ventas sin IVA', 'IVA', 'Tickets', 'Comensales', 'Observaciones'],
  // Columnas del registro de gastos
  GAS_COLS: ['Día', 'Fecha', 'Nº factura', 'Proveedor', 'Concepto', 'Categoría', 'Tipo',
             'Base (sin IVA)', 'IVA', 'Total', 'Forma de pago', 'Factura', 'Observaciones'],

  // Celdas clave del resumen de cada mes (las usa el Resumen anual y el panel)
  M: {
    ventas: 'D6', compras: 'D7', varInv: 'D8', cmv: 'D9', margenBruto: 'D10',
    personal: 'D11', explotacion: 'D12', resultado: 'D13',
    inversiones: 'D16', financiacion: 'D17', impuestos: 'D18', retirada: 'D19',
    cajaGenerada: 'D21',
    kFoodCost: 'D25', kPersonal: 'D26', kPrime: 'D27', kMargen: 'D28', kExplot: 'D29',
    saldoIni: 'L6', cobrado: 'L7', pagado: 'L8', saldoPrev: 'L9', efectivoReal: 'L10',
    bancosReal: 'L11', saldoReal: 'L12', diferencia: 'L13', pendiente: 'L14',
    cobEfectivo: 'L17', cobTarjeta: 'L18', cobBizum: 'L19', cobPlataformas: 'L20',
    invIni: 'L23', invFin: 'L24',
    dias: 'O6', tickets: 'O7', comensales: 'O8', ticketMedio: 'O9', gastoComensal: 'O10',
    ventaDia: 'O11', costesFijos: 'O14', margenContrib: 'O15', ventaNecesaria: 'O16',
    ventaNecDia: 'O17', difEquilibrio: 'O18',
    catFirst: 6, catLast: 33   // bloque "Gastos por categoría": F=categoría, G=tipo, H=€, I=%
  },

  // Configuración
  C: {
    nombre: 'B4', anio: 'B5', saldoIni: 'B6', invIni: 'B7', ivaVentas: 'B8',
    fcObj: 'B12', fcAlerta: 'C12', perObj: 'B13', perAlerta: 'C13',
    primeObj: 'B14', primeAlerta: 'C14', margenObj: 'B15', margenAlerta: 'C15',
    fcPlatoObj: 'B16', fcPlatoAlerta: 'C16',
    tpvDesde: 'B17',   // Datos del TPV desde (fecha): lo anterior se ignora al subir informes
    ingAuto: 'B18',    // Apuntar ingresos desde el TPV: Sí / No
    listFirst: 21,
    cat:   { col: 'A', last: 48 },  // Categoría de gasto | B Tipo | C Qué incluye
    tipos: { col: 'E', last: 27 },  // Tipo | F Qué significa
    canal: { col: 'H', last: 32 },  // Canal | I Tipo de canal
    catIng:{ col: 'K', last: 30 },  // Categoría de ingreso | L IVA
    pago:  { col: 'N', last: 28 },  // Formas de pago
    carta: { col: 'P', last: 40 },  // Categorías de la carta
    locales: { col: 'R', last: 30 } // Locales del grupo: R nombre | S enlace de su hoja
  },

  PL: { first: 5, last: 204 },     // Platos
  RE: { first: 5, last: 1504 },    // Recetas
  IN: { first: 5, last: 304 },     // Ingredientes
  TV: { first: 5, last: 2004 },    // Ventas TPV
  MAP: { first: 5, last: 504 }     // Mapeo nombres TPV → plato (columnas I:J de Ventas TPV)
};

/** Referencia absoluta a una celda de Configuración: CF('B5') -> 'Configuración'!$B$5 */
function CF(a1) { return "'" + L.CFG + "'!" + abs_(a1); }
/** Prefijo de hoja entre comillas */
function SH(name) { return "'" + name + "'!"; }
function abs_(a1) { return a1.replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2'); }
/** Rango de columna absoluto dentro de un bloque: col_('K', L.ING) -> $K$39:$K$188 */
function col_(c, blk) { return '$' + c + '$' + blk.first + ':$' + c + '$' + blk.last; }
/** Rango de una lista de Configuración: cfgList_('cat') -> 'Configuración'!$A$21:$A$48 */
function cfgList_(key, width) {
  var d = L.C[key], c1 = d.col, c2 = width ? String.fromCharCode(c1.charCodeAt(0) + width - 1) : c1;
  return "'" + L.CFG + "'!$" + c1 + '$' + L.C.listFirst + ':$' + c2 + '$' + d.last;
}
