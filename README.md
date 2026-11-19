# Registro de armarios eléctricos GPS · Versión 3.4

Aplicación web progresiva para registrar armarios eléctricos, posición GPS y fotografías localmente.

La versión 3.4 recupera automáticamente la carpeta raíz previamente autorizada y comprueba/crea la estructura `REGISTROS` y sus subcarpetas. Ya no muestra un selector manual de carpeta raíz. La cabecera muestra una fecha/hora fija de compilación: 19/11/2026 10:28.

Incluye previsualización de cada fotografía, selección de tipo (Ubicación, Interior o Documentación), y confirmación del alta del armario. El Service Worker utiliza una caché nueva `registro-armarios-gps-v3.4` para forzar la actualización de la aplicación instalada.

Consulta `LEEME.txt` para instalación, funcionamiento y publicación en GitHub Pages. El módulo `modArmarios.bas` se conserva sin cambios.
