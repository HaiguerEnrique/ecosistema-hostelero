---
name: actualizar
description: Actualiza el Ecosistema Hostelero de un local a la última versión del plugin, sin tocar sus datos ni sus personalizaciones. Úsala cuando el dueño diga «actualiza el Ecosistema», pregunte si hay versión nueva, o cuando la versión de ecosistema.json sea menor que la del núcleo del plugin.
---

# Actualizar el Ecosistema

La versión nueva viene dentro del plugin: `../../nucleo/` respecto a este SKILL.md (`VERSION` en `nucleo/Layout.js`). Para tener la última, el plugin tiene que estar al día (`/plugin` › marketplace `ecosistema` › actualizar) — compruébalo primero.

Los **datos** del dueño viven en su hoja y **no se tocan**. Sus **personalizaciones** son los archivos `app/Personal*` y tampoco se tocan. Lo que se sustituye es el resto de `app/` (el núcleo).

## Pasos
1. **Qué hay.** Lee `ecosistema.json` (versión instalada) y la del plugin. Si son iguales: «Ya tienes la última (X)». Si tiene varios locales, repite todo en cada carpeta.
2. **Novedades.** Compara los dos núcleos y cuéntale al dueño en 2–4 frases, sin jerga, qué cambia para él. Pregunta si seguimos.
3. **¿Alguien tocó el núcleo?** Calcula el sha256 de cada archivo de `app/` que no empiece por `Personal` y compáralo con `archivosNucleo` de `ecosistema.json`. Si alguno cambió, cópialo a `respaldo-<fecha>/`, díselo al dueño y ofrécele pasar ese cambio a un archivo `Personal*` (habilidad `personalizar`) antes de seguir.
4. **Sustituir.** En `app/`: borra los archivos que no empiecen por `Personal` y copia los de `../../nucleo/`. No toques `app/Personal*`.
5. **Subir.** `clasp push --force` y `clasp update-deployment <webDeploymentId> --description "Panel <versión>"` (así el panel del móvil también se actualiza).
6. **Terminar en la hoja (el dueño):** «Abre tu hoja y recarga la página». Al abrirse, la hoja ajusta sus pestañas sola si la versión lo necesita (sin tocar datos); puede tardar unos segundos. Si Google vuelve a pedir permisos, que los acepte.
7. **Guardar.** Actualiza `version`, `archivosNucleo` (nuevas huellas) y la fecha en `ecosistema.json`. Apunta en `NOTAS.md`: «<fecha> · actualizado a <versión>».

## Si algo falla
- Error al subir → nada ha cambiado en la hoja: corrige y repite desde el paso 5.
- La hoja da error tras actualizar → vuelve a copiar `respaldo`/el núcleo anterior si lo tienes, sube, y avisa al dueño para que lo comunique en la comunidad con el mensaje de error exacto.
