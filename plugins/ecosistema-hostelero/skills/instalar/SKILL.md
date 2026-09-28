---
name: instalar
description: Instala el Ecosistema Hostelero para un bar o restaurante — crea su hoja de control en su Google Drive, sube el programa, publica el panel del móvil y deja lista la carpeta de trabajo. Úsala cuando el usuario diga «instálame el Ecosistema», quiera empezar con la plantilla, añadir otro local, o no exista todavía una carpeta con ecosistema.json.
---

# Instalar el Ecosistema Hostelero

Vas a dejar funcionando, en el Google Drive del dueño, su hoja de control con el panel. Al final tendrá dos enlaces (la hoja y el panel del móvil) y una carpeta de trabajo en su ordenador.

**Cómo hablar:** el dueño es hostelero, no informático. Frases cortas, sin jerga, de tú. Explica cada paso en una línea *antes* de hacerlo y di qué tiene que hacer él cuando le toque. No le enseñes comandos salvo que pregunte. Calcula una hora en total.

**Archivos del plugin:** el núcleo (el programa de la hoja) está en `../../nucleo/` respecto a este SKILL.md, y las plantillas en `../../plantillas/`. Su versión es `VERSION` en `nucleo/Layout.js`.

## 0. Antes de nada: ¿ya está instalado?
Busca carpetas con `ecosistema.json` en `~/Ecosistema Hostelero/`. Si el local ya existe, no reinstales: ofrece `ecosistema-hostelero:actualizar` o `cargar-datos`.

## 1. Comprobar el ordenador
1. `node --version` → hace falta 18 o superior. Si no está, instálalo tú si puedes (Windows: `winget install OpenJS.NodeJS.LTS`; Mac: `brew install node`) o guíale a nodejs.org (versión LTS). Después, abre una terminal nueva.
2. `clasp --version` → si no está: `npm install -g @google/clasp`.

## 2. Conectar su Google
1. Pídele que abra https://script.google.com/home/usersettings con **la cuenta de Google del negocio** y active «API de Google Apps Script». Sin esto no se puede subir nada. Espera a que te confirme.
2. `clasp login` → se abre el navegador. Dile: «Elige la cuenta del negocio y pulsa Permitir. Es para que yo pueda crear la hoja en tu Drive; nadie más tiene acceso». Espera a que termine.
3. `clasp show-authorized-user` para confirmar la cuenta. Si no es la del negocio: `clasp logout` y repetir.

## 3. Datos mínimos
Pregunta (una cosa cada vez, y ofrece «no lo sé, lo dejamos para luego»):
- Nombre del local (y ciudad si tiene varios con el mismo nombre).
- Año con el que empieza (normalmente el actual).
- ¿Tiene más locales? (cada local es una instalación aparte; al final se enlazan).
- ¿Va a subir facturas otra persona sin Claude desde el móvil? (si sí, más adelante hará falta la clave de API).

## 4. Crear la carpeta y la hoja
1. Carpeta: `~/Ecosistema Hostelero/<Nombre del local>/` con `app/`, `datos/` y `NOTAS.md` (vacío con un título).
2. Dentro de esa carpeta: `clasp create-script --type sheets --title "<Nombre del local> · Control <año>" --rootDir app`.
   Crea la hoja en la raíz de su Drive y el `.clasp.json`. Si clasp dice «Project file already exists», hay un `.clasp.json` en una carpeta superior: créalo desde una carpeta temporal y mueve el `.clasp.json` aquí.
3. Edita `.clasp.json`: `"rootDir": "app"` y `"filePushOrder": ["app/Layout.js"]` (Layout.js tiene que ir el primero o la hoja se queda sin menú).
4. Copia **todos** los archivos de `../../nucleo/` a `app/` (sustituye el `appsscript.json` que creó clasp).
5. `clasp push --force`.
6. Panel del móvil: `clasp create-deployment --description "Panel"`. Apunta el id que devuelve (`AKfy…`). La web es `https://script.google.com/macros/s/<id>/exec`.
7. Escribe `ecosistema.json`:
   ```json
   { "local": "…", "anio": 2026, "version": "<VERSION del núcleo>", "instalado": "<AAAA-MM-DD>",
     "scriptId": "…", "hojaId": "<parentId de .clasp.json>", "hojaUrl": "https://docs.google.com/spreadsheets/d/<hojaId>/edit",
     "webDeploymentId": "AKfy…", "webUrl": "https://script.google.com/macros/s/AKfy…/exec",
     "archivosNucleo": { "<archivo>": "<sha256>", … } }
   ```
   `archivosNucleo` = huella sha256 de cada archivo copiado del núcleo (sirve para saber, al actualizar, si alguien los tocó).
8. Copia `../../plantillas/CLAUDE.md` a la carpeta rellenando `{{NOMBRE_LOCAL}}` y `{{FECHA}}`.

## 5. Construir la plantilla (lo hace el dueño)
Dile, paso a paso:
1. «Abre este enlace: <hojaUrl>». Espera unos segundos a que aparezca el menú **🍽️ Ecosistema** arriba.
2. «Pulsa 🍽️ Ecosistema › 🚀 Crear mi plantilla».
3. Google pedirá permisos y avisará de que la aplicación «no está verificada». Es normal: es su propia copia. «Configuración avanzada» → «Ir a…» → «Permitir».
4. Si tras permitir no pasa nada, que vuelva a pulsar «🚀 Crear mi plantilla». Se construye por partes: si Google corta a los pocos minutos, sale un aviso («vamos por la parte X») y hay que pulsar **🚀 Crear mi plantilla (seguir)** hasta que diga «Plantilla creada». Suelen ser 2 o 3 veces.
5. Al terminar, que recargue la página. Ya verá todas las pestañas y el menú completo.

## 6. Primeros datos
Prepara con `ecosistema-hostelero:cargar-datos` la configuración básica: nombre, año, dinero con el que empieza (caja + banco), inventario aproximado y, si tiene varios locales, la tabla de locales. Que lo cargue desde el menú. Después ofrece seguir con:
- sus Excel de cuentas, si los tiene;
- su carta y recetas (empezar por los 10 platos que más vende);
- los informes de su TPV (se suben en el panel, pestaña Ventas; primero los del año entero).

## 7. Cerrar
- Dale los dos enlaces: la hoja y el panel del móvil (que lo guarde en la pantalla de inicio; en el móvil entra con la misma cuenta de Google).
- Si otra persona va a subir facturas desde el móvil sin Claude: menú 🍽️ Ecosistema › 🔑 Configurar clave de IA (clave de console.anthropic.com; cuesta unos céntimos por factura) y compartirle la hoja.
- Varios locales: repite la instalación con cada uno y carga en cada hoja la tabla de locales con los enlaces de las otras.
- Apunta en `NOTAS.md` lo que hayas aprendido del negocio.

## Si algo falla
- «User has not enabled the Apps Script API» → paso 2.1, y esperar un par de minutos.
- El menú no aparece → recargar la hoja; si sigue sin salir, comprobar `filePushOrder` y volver a hacer `clasp push --force`.
- La hoja pide permisos cada vez → que los acepte con la misma cuenta con la que la creó.
