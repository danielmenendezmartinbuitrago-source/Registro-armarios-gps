# Registro de armarios eléctricos

Aplicación web en HTML para registrar armarios eléctricos con fecha, hora y posición GPS.

## Publicar con GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Descomprime este paquete en tu ordenador o teléfono.
3. Sube **`index.html`** y **`README.md`** a la carpeta raíz del repositorio (no subas el ZIP como único archivo).
4. En GitHub, entra en **Settings → Pages**.
5. En **Build and deployment**, selecciona **Deploy from a branch**.
6. Selecciona la rama `main` y la carpeta `/(root)`, y guarda.
7. Cuando termine la publicación, abre la dirección HTTPS que proporciona GitHub Pages.

## Importante sobre los datos y permisos

- GitHub Pages publica la aplicación, pero **no almacena ni sincroniza** los registros ni las fotografías.
- Los datos continúan guardándose en el almacenamiento local del dispositivo y en la carpeta que la aplicación tenga vinculada.
- Al abrir la aplicación desde la dirección de GitHub Pages, el navegador puede pedir que vuelvas a vincular la carpeta `ARMARIOS ELECTRICOS`. Esto es normal: los permisos guardados para otra dirección o para la apertura local no se transfieren automáticamente a la dirección de GitHub Pages.
- Usa Chrome en Android y concede acceso a la carpeta existente. No borres ni sustituyas los archivos que ya contiene.
- La aplicación necesita un contexto seguro (HTTPS) para las funciones de acceso a carpetas. GitHub Pages proporciona HTTPS.

## Archivo principal

- `index.html`: aplicación, versión 3.8.
