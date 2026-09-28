# Ecosistema Hostelero

Control financiero, escandallos y panel en el móvil para bares y restaurantes. Todo vive en **tu** Google Drive y lo instala y mantiene **tu propio Claude**.

- **Hoja de control** en Google Sheets: ventas, gastos, caja, coste de cada plato, punto de equilibrio.
- **Panel** con semáforos, en la hoja y en el móvil.
- **Claude** te instala todo, pasa tus Excel, calcula tus recetas, apunta tus facturas y te adapta la hoja.

## Qué necesitas
- Una cuenta de Google.
- Suscripción de Claude y la aplicación de Claude en el ordenador (con Claude Code).
- Opcional: una clave de la API de Claude, solo si otra persona sin Claude va a subir facturas desde el móvil.

## Instalar
En Claude Code:

```
/plugin marketplace add HaiguerEnrique/ecosistema-hostelero
/plugin install ecosistema-hostelero@ecosistema
```

Después escríbele a Claude: **«Instálame el Ecosistema»**.

## Qué sabe hacer tu Claude con el plugin
| Pídele | Qué hace |
| --- | --- |
| «Instálame el Ecosistema» | Crea tu hoja en tu Drive, el panel del móvil y tu carpeta de trabajo |
| «Pasa mis Excel a la hoja», «mete mi carta», «apúntame esta factura» | Prepara los datos; tú los revisas y los cargas con un botón |
| «Actualiza el Ecosistema» | Instala la versión nueva sin tocar tus datos ni tus cambios |
| «Añádeme…», «quiero que la hoja calcule…» | Adapta la hoja a tu negocio sin romper las actualizaciones |

## Actualizaciones
Cuando salga una versión nueva: `/plugin` → marketplace `ecosistema` → actualizar, y luego dile a Claude «actualiza el Ecosistema».

## Tus datos
Tu hoja, tus facturas y tus informes se quedan en tu Google Drive y en tu ordenador. Este repositorio solo contiene el programa.
