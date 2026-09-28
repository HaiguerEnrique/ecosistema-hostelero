# {{NOMBRE_LOCAL}} — Ecosistema Hostelero

Carpeta de trabajo del Ecosistema Hostelero de este local. La creó la habilidad `ecosistema-hostelero:instalar` el {{FECHA}}.

## Reglas para Claude
- Habla con el dueño en español de España, claro y sin tecnicismos. Trátale de tú.
- **No inventes cifras.** Los números salen de la hoja o de los documentos del dueño. Si falta un dato, pregúntalo.
- **El dueño confirma todo.** Tú preparas los datos; él los revisa y pulsa «Cargar» en la hoja.
- **No borres nada de la hoja ni de su Drive.** Si algo sobra, díselo y que lo borre él.
- **No toques los archivos del núcleo** (`app/` salvo `app/Personal*`): las actualizaciones los sustituyen. Los cambios a medida van en `app/Personal*.js` (habilidad `ecosistema-hostelero:personalizar`).
- Para subir el código: `clasp push --force` desde esta carpeta. Si la web del móvil tiene que ver el cambio: `clasp update-deployment <webDeploymentId de ecosistema.json> --description "Panel"`.

## Dónde está todo
- `ecosistema.json` — enlaces de la hoja y de la web del móvil, versión instalada, id del script.
- `app/` — código que va dentro de la hoja (núcleo + archivos `Personal*` del dueño).
- `datos/` — copias de los Excel, informes y facturas que el dueño ha pasado. Pueden tener datos personales: no salen de este ordenador.
- `NOTAS.md` — lo que Claude ha aprendido del negocio (cómo clasifica sus gastos, proveedores habituales, decisiones). Añade aquí lo importante al terminar cada tarea.

## Habilidades
- `ecosistema-hostelero:cargar-datos` — pasar Excel, carta, recetas o facturas a la hoja.
- `ecosistema-hostelero:actualizar` — instalar la versión nueva del Ecosistema.
- `ecosistema-hostelero:personalizar` — adaptar la hoja o el menú a este negocio.
