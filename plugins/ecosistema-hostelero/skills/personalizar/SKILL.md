---
name: personalizar
description: Adapta el Ecosistema Hostelero a un negocio concreto — categorías o canales propios, un informe o cálculo extra, una opción nueva en el menú de la hoja — sin romper las actualizaciones. Úsala cuando el dueño pida un cambio en su hoja o en su menú («añádeme…», «quiero que la hoja calcule…», «ponme un botón para…»).
---

# Personalizar el Ecosistema

Regla de oro: **el núcleo no se toca.** Todo lo que no sea `app/Personal*` se sustituye en cada actualización y cualquier cambio ahí se perdería.

## Qué se resuelve sin programar (primero esto)
Muchas peticiones son configuración, no código:
- Categorías de gasto, canales de venta, formas de pago, categorías de la carta, umbrales del semáforo, IVA → pestaña **Configuración** de la hoja. Explícale dónde o prepáralo con `cargar-datos`.
- Añadir platos, ingredientes o recetas → `cargar-datos`.
- Una pregunta sobre sus números → contéstala con los datos (panel o pregúntale la cifra); no hace falta cambiar nada.

## Cuando sí hace falta código
1. Crea o edita `app/Personal<Tema>.js` (por ejemplo `PersonalPropinas.js`). Un tema por archivo, con un comentario arriba que diga qué hace y cuándo se pidió.
2. Puedes usar lo del núcleo: `L` (mapa de celdas en `Layout.js`), `libro_()`, `cfg_()`, `gastos_(mes)`, `ingresos_(mes)`, `resumenMes_(mes)`.
3. **Menú propio:** define `function menuPersonal_(ui, menu) { menu.addSeparator().addItem('💶 Propinas del mes', 'propinasDelMes'); }`. El núcleo la llama al abrir la hoja. Solo puede haber una `menuPersonal_` en todos los `Personal*`.
4. **No redefinas funciones del núcleo** (mismo nombre): el resultado depende del orden de carga y se rompería al actualizar. Si lo que pide necesita cambiar cómo funciona el núcleo, dile que no se puede desde aquí y que lo proponga en la comunidad del Ecosistema.
5. No escribas en columnas con fórmulas ni cambies filas o columnas de la plantilla: el panel lee posiciones fijas.
6. `node --check` del archivo → `clasp push --force` → que el dueño recargue la hoja y lo pruebe. Si afecta al panel del móvil: `clasp update-deployment <webDeploymentId> --description "Panel"`.
7. Apunta en `NOTAS.md` qué se personalizó y por qué.
