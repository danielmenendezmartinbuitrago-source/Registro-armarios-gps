# Registro de armarios eléctricos GPS · Versión 3.0

**Versión:** 3.0  
**Fecha de actualización:** 18/11/2026

Aplicación web estática instalable como PWA para registrar armarios eléctricos con referencia, planta, observaciones de catálogo, fecha/hora y posición GPS.

## Publicación en GitHub Pages

Sube **el contenido de esta carpeta** a la raíz del repositorio. En **Settings → Pages**, selecciona **Deploy from a branch**, rama `main` y carpeta `/(root)`. La aplicación necesita HTTPS para utilizar la geolocalización.

## Funciones

- Referencias con formato `PREFIJO-XX-XX`.
- Medición GPS de 30, 60, 90 o 120 segundos, guardando la lectura de mayor precisión.
- Catálogo TXT de tres columnas separadas por tabuladores: referencia actual, observación y referencia antigua.
- Búsqueda parcial de la referencia antigua. Al seleccionar un resultado se recuperan la referencia actual y sus observaciones.
- Entrada exclusivamente manual: se han eliminado los controles y funciones de dictado por voz.
- Captura de varias fotografías por armario desde la cámara del móvil.
- Selección de la carpeta física `REGISTROS` del dispositivo Android (Chrome 132 o posterior).
- Registros guardados en `REGISTRO_ARMARIOS.json` y catálogo en `CATALOGO_ARMARIOS.txt`.
- Fotografías guardadas como archivos JPEG en `FOTOGRAFIAS/PREFIJO/`, con nombres correlativos `PREFIJO-XX-XX_01.jpg`, `_02.jpg`, etc.
- Creación automática de subcarpetas para F01, F02, F03, F03B, F04, F04B, F05, F06, F07, F08, F09, F10, F11, F12, F13, F14, F15, F16, F16B, F17, F18, F19, F19B, F20, F41 y AMR.
- Exportación Excel (`.xlsx`) y TXT delimitado por `;`, con una fila por registro e incluyendo los nombres de las fotografías.
- Consulta y filtrado de los registros guardados en el dispositivo.

## Datos y privacidad

Los registros, el catálogo y las fotografías se guardan en la carpeta `REGISTROS` que selecciona el usuario. Chrome conserva el identificador de la carpeta y puede volver a solicitar permiso al abrir la aplicación. No se envían datos a GitHub ni a un servidor, y no se sincronizan entre dispositivos. La primera conexión intenta migrar los registros y fotografías antiguos de localStorage/IndexedDB al directorio seleccionado. Las exportaciones Excel y TXT incluyen fecha y hora, referencia actual, referencia antigua, planta, latitud, longitud, precisión GPS, observaciones y nombres de las fotografías. Los archivos de imagen no se incluyen en las exportaciones y no se sincronizan entre dispositivos.

El catálogo se selecciona desde el formulario. La aplicación web solicita permiso explícito para escribir en la carpeta `REGISTROS` mediante File System Access API. Esta función requiere un navegador Chromium compatible y HTTPS.

## Archivos

- `index.html`: aplicación (identificada como versión 3.0).
- `manifest.webmanifest`: configuración PWA.
- `sw.js`: Service Worker y caché de la versión 3.0.
- `icon-192.png` y `icon-512.png`: iconos.
- `.nojekyll`: configuración de GitHub Pages.
- `modArmarios.bas`: módulo VBA para la integración de escritorio.
