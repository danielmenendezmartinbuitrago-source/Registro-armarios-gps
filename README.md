# Registro de armarios eléctricos

Aplicación web estática para registrar armarios eléctricos con referencia, planta, observaciones, fotografías (si se incorporan en una futura versión), fecha/hora y posición GPS.

## Publicación en GitHub Pages

Este repositorio está preparado para el repositorio `registro-armarios-gps`.

1. Sube **el contenido de esta carpeta** a la raíz del repositorio, no la carpeta contenedora.
2. En GitHub abre **Settings → Pages**.
3. En **Build and deployment**, selecciona **Deploy from a branch**.
4. Selecciona la rama `main` y la carpeta `/(root)`, y guarda.
5. Activa HTTPS si GitHub no lo activa automáticamente.

La dirección prevista del sitio es:
https://danielmenendezmartinbuitrago-source.github.io/registro-armarios-gps/

## Archivos

- `index.html`: aplicación.
- `manifest.webmanifest`: configuración instalable (PWA).
- `sw.js`: caché básica para uso sin conexión después de la primera carga.
- `icon-192.png` y `icon-512.png`: iconos de la aplicación.
- `.nojekyll`: evita el procesamiento de GitHub Pages con Jekyll.

## Datos y privacidad

Los registros se guardan en el almacenamiento local (`localStorage`) del navegador/dispositivo. Las fotografías se guardan en IndexedDB en el mismo dispositivo y reciben nombres correlativos basados en la referencia (por ejemplo, `F01-12-05_01.jpg`). No se envían automáticamente a GitHub ni a un servidor y no se sincronizan entre dispositivos. Las exportaciones Excel/TXT incluyen los nombres de las fotografías, no los archivos de imagen. Conserva una copia segura de las imágenes si necesitas trasladarlas a otro dispositivo.

El catálogo de observaciones se carga desde un archivo TXT en el propio formulario. Su formato es de tres columnas separadas por tabuladores (TAB): referencia actual, observación y referencia antigua.

El TXT de exportación de registros usa una fila por armario y `;` como separador de campos.

## Permisos

- La geolocalización requiere HTTPS y permiso de ubicación del navegador.
- El dictado por voz depende del soporte del navegador y de sus permisos. Se recomienda probarlo en Chrome.
- El modo sin conexión se activa después de la primera visita con conexión.
