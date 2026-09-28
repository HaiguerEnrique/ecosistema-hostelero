---
name: cargar-datos
description: Pasa datos del hostelero a su hoja del Ecosistema Hostelero — sus Excel o cuentas antiguas, la configuración inicial, la carta con precios, las recetas (escandallos), los ingredientes y las facturas de proveedores. Úsala cuando diga «pasa mis Excel», «apúntame esta factura», «mete mi carta», «calcula lo que me cuesta este plato» o tras instalar. Requiere una carpeta con ecosistema.json.
---

# Cargar datos en la hoja

Tú **preparas** los datos; el dueño los **revisa y carga** con un botón. Nunca se escribe nada en su hoja sin que él pulse «Sí».

**Cómo funciona:** escribes `app/PersonalDatos.js` con `var DATOS_CLAUDE = {…}`, lo subes con `clasp push --force` y el dueño pulsa en su hoja **🍽️ Ecosistema › 📥 Cargar datos preparados por Claude**. La hoja comprueba todo, le enseña un resumen y, si dice que sí, lo escribe. Después retiras el archivo.

**Formato exacto:** está al principio de `app/Importar.js` (léelo antes de empezar). Categorías, canales, formas de pago y categorías de carta que trae la hoja de fábrica: en `app/Construir.js` (`construirConfiguracion_`). Las que el negocio haya añadido están apuntadas en `NOTAS.md`.

**Cómo hablar:** sin tecnicismos. Enséñale lo que vas a cargar como lo vería él (tabla corta, cifras redondas, totales por mes) y pregúntale lo que no esté claro. Nunca rellenes un hueco inventando.

## Pasos
1. **Recoger.** Pídele los archivos (Excel, PDF, fotos) y guárdalos en `datos/`. Si son del TPV (informes de ventas o de facturas del día), no pasan por aquí: se suben en el panel, pestaña Ventas.
2. **Entender y proponer.** Lee los archivos y propón cómo se colocan:
   - Cada gasto a una **categoría de gasto** de su hoja. Las compras mixtas (cash & carry) se reparten por líneas: comida, bebida, envases, limpieza…
   - La **retirada del dueño**, los **préstamos**, las **inversiones** (maquinaria, reformas) y los **impuestos** van a sus categorías, no a gastos normales: si no, el resultado del mes sale mal.
   - Ventas: por día si las tiene; si solo tiene totales del mes, una línea por mes y canal (sin día).
   - Enséñale el reparto (totales por categoría y mes) y **espera su visto bueno**. Si algo no encaja con ninguna categoría, pregúntale si crear una nueva (`categoriasGasto`).
3. **Escribir `app/PersonalDatos.js`.** Con un `id` nuevo y claro (`excel-2026-ene-sep`, `factura-makro-2026-10-03`…) y una `descripcion` en su idioma. Importes en euros con 2 decimales; porcentajes como fracción (10 % = 0.1); meses 1–12.
   - Comprueba antes de subir: `node --check app/PersonalDatos.js` y que los totales por mes cuadran con sus archivos.
   - Mucho volumen (más de ~3.000 líneas o ~500 KB): por tandas, un archivo cada vez, cada uno con su `id`.
4. **Subir:** `clasp push --force`.
5. **Cargar (el dueño):** «Abre tu hoja y pulsa 🍽️ Ecosistema › 📥 Cargar datos preparados por Claude. Revisa el resumen y dale a Sí». Si la hoja da un error, te lo leerá: corrígelo y vuelve al paso 3.
6. **Retirar:** cuando te diga que se ha cargado, deja `app/PersonalDatos.js` con una sola línea, `var DATOS_CLAUDE = null;`, y `clasp push --force`. Los datos ya están en la hoja; el archivo solo era el transporte. (No lo borres: si solo desaparece un archivo, clasp responde «Script is already up to date» y no sube nada, así que los datos seguirían dentro.)
7. **Apuntar** en `NOTAS.md` las decisiones (cómo se clasifica cada proveedor, categorías nuevas, platos dados de alta).

## Todo de una vez
Si en la misma conversación el dueño te da varias cosas (ventas, escandallos, facturas, su carta…), júntalas en **un solo** `PersonalDatos.js`: la hoja lo comprueba todo junto, enseña un único resumen y se carga con un solo «Sí». El orden lo resuelve la hoja (primero configuración y listas, luego ingredientes y platos, luego recetas y ventas). Díselo así: «Pásamelo todo y te lo dejo listo para cargar de una vez».

## Casos frecuentes
- **Configuración inicial:** `configuracion` (nombre, año, saldo inicial = caja + bancos el día que empieza, inventario inicial, IVA de ventas si no es el 10 %) y `locales` si tiene varios.
- **Carta y recetas:** `platos` (nombre, categoría de la carta, precio con IVA) + `ingredientes` (unidad kg, l o ud; precio del formato sin IVA y cuánto trae el formato; merma si se pierde al limpiar o cocinar) + `recetas` (cantidades en la unidad del ingrediente: 150 g de carne = 0.150 kg). Los **gramos son del plato servido**: si el ingrediente se pierde al cocinar (patatas, carne desmechada), pon la merma en el ingrediente. Empieza por los 10 platos que más vende.
- **Una factura (sin clave de API):** lee la foto o el PDF, propón las líneas (proveedor, número, fecha, base, IVA y categoría por línea; reparte las compras mixtas) y cuadra base + IVA con el total. Cárgala como `gastos` con `observaciones: "Factura <proveedor> <número>"`. Recuérdale que guarde la factura en su Drive, en `Facturas/<año>/<trimestre>` (menú 📁 Configurar carpetas de Drive la crea).
- **Ventas por plato de un TPV que no es Foodyservice:** del informe de «ventas por producto» de su programa, `ventasPlatos` con una línea por producto y mes (`nombre` tal cual sale en el TPV, `unidades`, `importe` con IVA y, si lo trae, `udsLlevar`). Si el nombre no es igual que el del plato de la carta, pon `plato` y la hoja recordará la relación. Así funcionan el food cost real y la clasificación de la carta. (Los informes de Foodyservice no pasan por aquí: se suben en el panel.)
- **Productos del TPV «sin plato asignado»** (el panel los lista en Ventas): suele ser un nombre cambiado en el TPV o algo que no está en la carta (salsas, envases). Propón la relación con `nombresTpv: [['NOMBRE TPV', 'Plato']]` y, si falta el plato, créalo en `platos` (con su receta si se sabe) en la misma carga.
- **Cierre de mes:** `cierres` con el efectivo y el banco del último día y el inventario final.
- **Ya cargado:** si la hoja avisa de que ese `id` ya se cargó, NO sigas salvo que el dueño quiera duplicarlo a propósito.
