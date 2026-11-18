# Registro de armarios eléctricos GPS · Versión 2.7

**Versión:** 2.7  
**Fecha de actualización:** 18/11/2026

Aplicación web estática instalable como PWA para registrar armarios eléctricos con referencia, planta, observaciones de catálogo, fecha/hora y posición GPS.

## Publicación en GitHub Pages

Sube **el contenido de esta carpeta** a la raíz del repositorio. En **Settings → Pages**, selecciona **Deploy from a branch**, rama `main` y carpeta `/(root)`. La aplicación necesita HTTPS para utilizar la geolocalización.

## Funciones

- Referencias con formato `PREFIJO-XX-XX`.
- Medición GPS de 30, 60, 90 o 120 segundos, guardando la lectura de mayor precisión.
- Catálogo TXT de tres columnas separadas por tabuladores: referencia actual, observación y referencia antigua.
- Búsqueda parcial y por voz de la referencia antigua. Al seleccionar un resultado se recuperan la referencia actual y sus observaciones.
- Captura de varias fotografías por armario desde la cámara del móvil.
- Fotografías guardadas localmente en IndexedDB con nombres correlativos `PREFIJO-XX-XX_01.jpg`, `_02.jpg`, etc.
- Exportación Excel (`.xlsx`) y TXT delimitado por `;`, con una fila por registro e incluyendo los nombres de las fotografías.
- Consulta y filtrado de los registros guardados en el dispositivo.

## Datos y privacidad

Los registros y el catálogo se guardan en el almacenamiento local del navegador (`localStorage`). Las fotografías se guardan localmente en IndexedDB. No se envían automáticamente a GitHub ni a un servidor, y no se sincronizan entre dispositivos. Las exportaciones Excel y TXT incluyen fecha y hora, referencia actual, referencia antigua, planta, latitud, longitud, precisión GPS, observaciones y nombres de las fotografías. Los archivos de imagen no se incluyen en las exportaciones y no se sincronizan entre dispositivos.

El catálogo se selecciona desde el formulario. La aplicación web no puede escribir directamente en las carpetas locales `ARMARIOS ELECTRICOS/REGISTROS` por las restricciones de seguridad del navegador. Esa estructura corresponde a la integración de escritorio Excel/VBA documentada en `LEEME_INTEGRACION.txt`.

## Archivos

- `index.html`: aplicación (identificada como versión 2.7).
- `manifest.webmanifest`: configuración PWA.
- `sw.js`: Service Worker y caché de la versión 2.7.
- `icon-192.png` y `icon-512.png`: iconos.
- `.nojekyll`: configuración de GitHub Pages.
- `modArmarios.bas`: módulo VBA para la integración de escritorio.
