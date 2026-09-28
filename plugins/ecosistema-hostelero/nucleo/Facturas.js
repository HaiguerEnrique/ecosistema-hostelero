/**
 * Facturas: la IA lee la factura (foto o PDF), propone cómo apuntarla y, cuando el dueño
 * confirma, se guarda en Facturas/<año>/<trimestre> y se añade al registro de gastos del mes.
 */

var MAX_BYTES_FACTURA = 15 * 1024 * 1024;

/** Paso 1: sube el archivo a "Facturas/_por revisar" y pide a la IA que lo lea. */
function leerFactura(archivo, opciones) {
  opciones = opciones || {};
  var bytes = Utilities.base64Decode(archivo.base64);
  if (bytes.length > MAX_BYTES_FACTURA) throw new Error('El archivo pesa demasiado (máx. 15 MB).');
  var mime = archivo.mime;
  if (!/^(application\/pdf|image\/(jpeg|png|webp|gif))$/.test(mime)) throw new Error('Formato no admitido. Usa PDF, JPG o PNG.');

  var c = carpetas_(true);
  var revisar = sub_(c.base, 'Facturas', true);
  revisar = sub_(revisar, '_por revisar', true);
  var file = revisar.createFile(Utilities.newBlob(bytes, mime, archivo.nombre || 'factura'));

  var cfg = cfg_();
  var cats = cfg.categorias.map(function (x) { return x.categoria; });
  var bloque = mime === 'application/pdf'
    ? { type: 'document', source: { type: 'base64', media_type: mime, data: archivo.base64 } }
    : { type: 'image', source: { type: 'base64', media_type: mime, data: archivo.base64 } };
  var desglosar = opciones.desglosar !== false;

  var esquema = {
    type: 'object', additionalProperties: false,
    required: ['es_factura', 'proveedor', 'nif_proveedor', 'numero', 'fecha', 'lineas', 'total', 'forma_pago', 'notas'],
    properties: {
      es_factura: { type: 'boolean' },
      proveedor: { type: 'string' },
      nif_proveedor: { type: 'string' },
      numero: { type: 'string' },
      fecha: { type: 'string', description: 'AAAA-MM-DD, o vacío si no se ve' },
      lineas: { type: 'array', items: { type: 'object', additionalProperties: false,
        required: ['categoria', 'concepto', 'base', 'iva'],
        properties: { categoria: { type: 'string', enum: cats }, concepto: { type: 'string' },
                      base: { type: 'number' }, iva: { type: 'number' } } } },
      total: { type: 'number' },
      forma_pago: { type: 'string', enum: [''].concat(cfg.formasPago) },
      notas: { type: 'string' }
    }
  };
  var instrucciones = [
    'Lee esta factura o ticket de un proveedor de un negocio de hostelería en España y prepárala para apuntarla en su control de gastos.',
    '- proveedor: nombre comercial corto (ej. "Makro", "Endesa"). numero: número de factura tal cual. fecha: fecha de la factura.',
    '- lineas: importes SIN IVA (base) y la cuota de IVA en euros, agrupados por categoría de gasto. La suma de base + iva de todas las líneas debe cuadrar con el total de la factura; si hay recargo de equivalencia, portes o redondeos, inclúyelos en la línea que corresponda.',
    desglosar
      ? '- Si una factura de mercancía mezcla comida, bebida, envases y limpieza, sepárala en una línea por categoría ("Compras comida", "Compras bebida", "Envases y desechables", "Limpieza e higiene"...) sumando los artículos de cada una. Si no se puede distinguir, usa "Compras de mercancía (sin desglosar)".'
      : '- Las facturas de mercancía van en una sola línea "Compras de mercancía (sin desglosar)".',
    '- concepto: descripción corta de qué es (ej. "Compra semanal", "Luz agosto").',
    '- forma_pago: solo si aparece en la factura (tarjeta, efectivo, transferencia, domiciliación); si no, vacío.',
    '- es_factura: false si el documento no es una factura o ticket de gasto.',
    '- notas: cualquier cosa dudosa que el dueño deba revisar (importes ilegibles, fecha rara, IVA que no cuadra). Vacío si todo está claro.'
  ].join('\n');

  var resp;
  try {
    resp = claude_({
      max_tokens: 8000,
      output_config: { effort: 'medium', format: { type: 'json_schema', schema: esquema } },
      messages: [{ role: 'user', content: [bloque, { type: 'text', text: instrucciones }] }]
    });
  } catch (e) { file.setTrashed(true); throw e; }
  var datos = JSON.parse(textoDe_(resp));
  datos.archivoId = file.getId();
  datos.archivoUrl = file.getUrl();
  datos.duplicado = buscarDuplicado_(datos.proveedor, datos.numero);
  datos.anioPlantilla = cfg.anio;
  var suma = datos.lineas.reduce(function (s, l) { return s + l.base + l.iva; }, 0);
  datos.descuadre = Math.round((datos.total - suma) * 100) / 100;
  return datos;
}

/** Paso 2: el dueño ha revisado (y quizá corregido) la propuesta. Guarda archivo y filas. */
function guardarFactura(p) {
  var cfg = cfg_();
  var m = String(p.fecha || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Falta la fecha de la factura (AAAA-MM-DD).');
  var anio = +m[1], mesIdx = +m[2] - 1, dia = +m[3];
  if (anio !== cfg.anio) throw new Error('La factura es de ' + anio + ' y esta hoja es de ' + cfg.anio + '. Apúntala en el archivo de ese año.');
  if (!p.lineas || !p.lineas.length) throw new Error('La factura no tiene líneas.');

  var fecha = new Date(anio, mesIdx, dia);
  var destino = carpetaTrimestre_(fecha);
  var file = DriveApp.getFileById(p.archivoId);
  var ext = (file.getName().match(/\.[a-z0-9]+$/i) || [''])[0] || (file.getMimeType() === 'application/pdf' ? '.pdf' : '.jpg');
  var limpio = function (s) { return String(s || '').replace(/[\\/:*?"<>|]/g, '').trim(); };
  file.setName([p.fecha, limpio(p.proveedor) || 'Proveedor', limpio(p.numero)].filter(String).join(' · ') + ext);
  file.moveTo(destino);

  var sh = libro_().getSheetByName(L.MESES[mesIdx]);
  var G = L.GAS, n = G.last - G.first + 1;
  var vals = sh.getRange(G.first, 1, n, 13).getValues();
  var libre = -1;
  for (var i = 0; i < n; i++) {
    var f = vals[i];
    var vacia = [0, 2, 3, 4, 5, 7, 8, 10, 11, 12].every(function (c) { return f[c] === ''; });
    if (vacia) {
      var hueco = true;
      for (var j = 1; j < p.lineas.length; j++) {
        var g = vals[i + j];
        if (!g || ![0, 3, 4, 5, 7].every(function (c) { return g[c] === ''; })) { hueco = false; break; }
      }
      if (hueco) { libre = i; break; }
    }
  }
  if (libre < 0) throw new Error('El registro de gastos de ' + L.MESES[mesIdx] + ' está lleno.');

  var fila = G.first + libre;
  p.lineas.forEach(function (l, k) {
    var r = fila + k;
    sh.getRange(r, 1).setValue(dia);
    sh.getRange(r, 3, 1, 4).setValues([[p.numero || '', p.proveedor || '', l.concepto || '', l.categoria]]);
    sh.getRange(r, 8, 1, 2).setValues([[Math.round(l.base * 100) / 100, Math.round(l.iva * 100) / 100]]);
    sh.getRange(r, 11).setValue(p.forma_pago || '');
    sh.getRange(r, 12).setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Ver factura').setLinkUrl(file.getUrl()).build());
    if (p.notas) sh.getRange(r, 13).setValue(p.notas);
  });
  SpreadsheetApp.flush();
  return { mes: L.MESES[mesIdx], fila: fila, lineas: p.lineas.length, carpeta: destino.getName(), url: file.getUrl() };
}

function descartarFactura(archivoId) {
  try { DriveApp.getFileById(archivoId).setTrashed(true); } catch (e) {}
  return true;
}

function buscarDuplicado_(proveedor, numero) {
  if (!numero) return null;
  var n = String(numero).trim().toLowerCase(), p = String(proveedor || '').trim().toLowerCase();
  var hit = gastos_('todos').filter(function (f) {
    return String(f.factura).trim().toLowerCase() === n &&
      (!p || String(f.proveedor).toLowerCase().indexOf(p) >= 0 || p.indexOf(String(f.proveedor).toLowerCase()) >= 0);
  })[0];
  return hit ? { mes: hit.mes, fila: hit.fila } : null;
}
