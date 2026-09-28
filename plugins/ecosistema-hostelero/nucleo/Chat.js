/**
 * Asistente del panel: responde preguntas sobre el negocio con los datos reales de la hoja.
 * La IA no ve la hoja entera: pide lo que necesita con herramientas y el código se lo da ya calculado.
 */

/** Se construyen al usarse: L (Layout.js) puede no estar cargado todavía al leer este archivo. */
function herramientas_() {
  return [
  { name: 'resumen_anual',
    description: 'Cuenta de resultados, indicadores y caja de los 12 meses del año, más los totales anuales. Úsala para preguntas del año, comparar meses o ver tendencias.',
    input_schema: { type: 'object', properties: {}, additionalProperties: false } },
  { name: 'resumen_mes',
    description: 'Resumen completo de un mes: ventas, coste de mercancía, personal, explotación, resultado, indicadores con su estado, caja, cobros por forma de pago, datos operativos, punto de equilibrio y gastos por categoría.',
    input_schema: { type: 'object', additionalProperties: false, required: ['mes'],
      properties: { mes: { type: 'string', enum: L.MESES } } } },
  { name: 'gastos',
    description: 'Registro de gastos y facturas. Puede devolver las líneas (filtradas por categoría o proveedor) o totales agrupados. Importes sin IVA en "base". Para "cuánto gasto en X" o "mis mayores proveedores" usa agrupar.',
    input_schema: { type: 'object', additionalProperties: false, required: ['mes'],
      properties: {
        mes: { type: 'string', enum: L.MESES.concat(['todos']), description: 'Mes o "todos" para el año completo' },
        categoria: { type: 'string', description: 'Filtra por categoría exacta (opcional)' },
        proveedor: { type: 'string', description: 'Filtra por texto contenido en el proveedor o el concepto (opcional)' },
        agrupar_por: { type: 'string', enum: ['nada', 'categoria', 'proveedor', 'mes', 'tipo'], description: 'Devuelve totales agrupados en lugar de líneas' } } } },
  { name: 'ingresos',
    description: 'Registro de ingresos: ventas por día y canal, con cobros en efectivo, tarjeta, Bizum y plataformas. Puede agrupar por canal, día o mes.',
    input_schema: { type: 'object', additionalProperties: false, required: ['mes'],
      properties: {
        mes: { type: 'string', enum: L.MESES.concat(['todos']) },
        agrupar_por: { type: 'string', enum: ['nada', 'canal', 'dia', 'mes', 'categoria'] } } } },
  { name: 'ventas_tpv',
    description: 'Datos del TPV de un mes: ventas por noche, pedidos, ticket medio, canales (sala, llevar, domicilio), media por día de la semana, ventas por hora, unidades e importe de cada plato con su margen, camareros con % de cancelaciones, y CMV teórico (según escandallos) frente al real con el descuadre en euros. Vacío si no se han subido los informes del TPV de ese mes.',
    input_schema: { type: 'object', additionalProperties: false, required: ['mes'],
      properties: { mes: { type: 'string', enum: L.MESES } } } },
  { name: 'carta',
    description: 'Platos de la carta con PVP, coste de la ración, food cost %, margen por plato, unidades vendidas, margen total y clasificación (estrella, caballo, enigma, perro). Vacía si aún no se han cargado escandallos.',
    input_schema: { type: 'object', properties: {}, additionalProperties: false } },
  { name: 'configuracion',
    description: 'Datos del negocio, referencias de indicadores (objetivo y alerta), categorías de gasto con su tipo, canales de venta y formas de pago.',
    input_schema: { type: 'object', properties: {}, additionalProperties: false } }
  ];
}

function ejecutarHerramienta_(nombre, e) {
  switch (nombre) {
    case 'resumen_anual': {
      var a = resumenAnual_();
      a.meses = a.meses.filter(function (m) { return m.conDatos; }).map(function (m) {
        return { mes: m.mes, ventas: m.ventas, cmv: m.cmv, personal: m.personal, explotacion: m.explotacion,
                 resultado: m.resultado, cajaGenerada: m.cajaGenerada, retirada: m.retirada, kpi: m.kpi,
                 dias: m.dias, ventaNecesaria: m.ventaNecesaria, pendiente: m.pendiente };
      });
      return a;
    }
    case 'resumen_mes': return resumenMes_(e.mes);
    case 'gastos': {
      var g = gastos_(e.mes);
      if (e.categoria) g = g.filter(function (f) { return String(f.categoria).toLowerCase() === String(e.categoria).toLowerCase(); });
      if (e.proveedor) {
        var t = String(e.proveedor).toLowerCase();
        g = g.filter(function (f) { return (String(f.proveedor) + ' ' + String(f.concepto)).toLowerCase().indexOf(t) >= 0; });
      }
      if (e.agrupar_por && e.agrupar_por !== 'nada') return { agrupado_por: e.agrupar_por, grupos: agrupar_(g, e.agrupar_por, 'base'), lineas_totales: g.length };
      return { lineas: g.slice(0, 250).map(function (f) { delete f.fila; return f; }), total_lineas: g.length,
               nota: g.length > 250 ? 'Se muestran 250 de ' + g.length + '. Usa agrupar_por o filtros.' : '' };
    }
    case 'ingresos': {
      var i = ingresos_(e.mes);
      if (e.agrupar_por && e.agrupar_por !== 'nada') return { agrupado_por: e.agrupar_por, grupos: agrupar_(i, e.agrupar_por, 'sinIva'), nota: 'Importes sin IVA' };
      return { lineas: i.slice(0, 250), total_lineas: i.length };
    }
    case 'ventas_tpv': return ventasTpvMes_(e.mes);
    case 'carta': return { platos: platos_() };
    case 'configuracion': return cfg_();
  }
  throw new Error('Herramienta desconocida: ' + nombre);
}

function sistemaChat_(locs, elegido) {
  var c = cfg_();
  var varios = locs && locs.length > 1;
  return [
    varios
      ? 'Eres el asistente financiero de un grupo de hostelería con ' + locs.length + ' locales: ' + locs.map(function (l) { return l.nombre; }).join(', ') + '. Cada local tiene su propia hoja; en cada herramienta indica el local con el parámetro "local" (si la pregunta es de todos o compara locales, consulta cada uno por separado). ' +
        (elegido ? 'El dueño está mirando ahora el local ' + elegido + ': úsalo si no dice otro.' : 'El dueño está mirando todos los locales a la vez.') +
        ' Hablas con el dueño en español, de tú, claro y directo, como un asesor de confianza que conoce la hostelería por dentro.'
      : 'Eres el asistente financiero de ' + c.nombre + ', un negocio de hostelería. Hablas con su dueño en español, de tú, claro y directo, como un asesor de confianza que conoce la hostelería por dentro.',
    'Tus datos salen de su hoja de control financiero del año ' + c.anio + '. Antes de dar cualquier cifra, consúltala con las herramientas: nunca inventes ni estimes un número que no venga de ellas. Si el dato no está registrado, dilo y explica qué tendría que apuntar para tenerlo.',
    'Cómo está organizada la hoja: importes sin IVA salvo que se diga lo contrario. Resultado operativo = ventas − coste de mercancía − personal − gastos de explotación. Inversiones, cuotas de préstamos, impuestos y la retirada del propietario salen de caja pero no son gasto del negocio y no restan del resultado. Food cost = coste de mercancía / ventas; si no se hace inventario, el coste de mercancía es igual a las compras. Prime cost = (mercancía + personal) / ventas. Las compras mixtas (tipo Makro) pueden estar en "Compras de mercancía (sin desglosar)". El CMV teórico sale de los escandallos y de las ventas del TPV; si la cobertura es baja o faltan recetas, dilo. Un mes a medias puede tener gastos aún sin apuntar (por ejemplo, nóminas pendientes): tenlo en cuenta antes de sacar conclusiones y avísalo.',
    'Al responder: primero la respuesta en una o dos frases, luego el detalle que la sostiene. Cuando recomiendes algo, di qué dato te lleva a ello. Sé breve. Formato simple: párrafos cortos, listas con guiones y **negritas** para las cifras clave. Nada de tablas anchas. Importes con formato español (1.234,56 €).',
    'Hoy es ' + Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy') + '.'
  ].join('\n\n');
}

/**
 * Turno de chat. historial: [{role:'user'|'assistant', text}] (solo texto de turnos anteriores).
 * Devuelve { texto, herramientas: [nombres usados] }.
 */
function preguntar(historial, pregunta, localId) {
  var locs = locales_(), varios = locs.length > 1;
  var elegido = localId && localId !== 'todos' ? locs.filter(function (l) { return l.id === localId; })[0] : (varios ? null : locs[0]);
  var porNombre = {}; locs.forEach(function (l) { porNombre[l.nombre] = l.id; });
  var tools = herramientas_();
  if (varios) tools.forEach(function (t) {
    t.input_schema.properties.local = { type: 'string', enum: locs.map(function (l) { return l.nombre; }), description: 'Local al que se refiere la consulta' };
    t.input_schema.required = (t.input_schema.required || []).concat(['local']);
  });
  var mensajes = (historial || []).slice(-12).map(function (m) { return { role: m.role, content: m.text }; });
  mensajes.push({ role: 'user', content: pregunta });
  var usadas = [];
  var ejecutar = function (nombre, e) {
    var id = e.local ? porNombre[e.local] : (elegido ? elegido.id : null);
    delete e.local;
    return enLocal_(id, function () { return ejecutarHerramienta_(nombre, e); });
  };
  for (var paso = 0; paso < 10; paso++) {
    var resp = enLocal_(elegido ? elegido.id : null, function () { return claude_({
      max_tokens: 8000,
      system: sistemaChat_(locs, elegido ? elegido.nombre : null),
      tools: tools,
      output_config: { effort: 'medium' },
      messages: mensajes
    }); });
    if (resp.stop_reason !== 'tool_use') return { texto: textoDe_(resp) || '(sin respuesta)', herramientas: usadas };
    mensajes.push({ role: 'assistant', content: resp.content });
    var resultados = resp.content.filter(function (b) { return b.type === 'tool_use'; }).map(function (b) {
      usadas.push(b.name);
      try {
        return { type: 'tool_result', tool_use_id: b.id, content: JSON.stringify(ejecutar(b.name, b.input || {})) };
      } catch (err) {
        return { type: 'tool_result', tool_use_id: b.id, content: 'Error: ' + err.message, is_error: true };
      }
    });
    mensajes.push({ role: 'user', content: resultados });
  }
  return { texto: 'La consulta necesitaba demasiados pasos. Prueba a preguntar algo más concreto.', herramientas: usadas };
}
