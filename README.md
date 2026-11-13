# Dashboard Inspección de contenedores (GitHub Pages)

Página estática. En GitHub solo está el código; los datos se leen en vivo desde SharePoint con la cuenta Microsoft de quien abre la página. Los flujos de Power Automate ya actualizan los CSV cada día a las 10 pm, por lo que la página siempre muestra lo último, sin tareas programadas en GitHub.

Archivos: `index.html`, `shim.js`, `config.js` (el único que se edita).

## 1. Registrar la aplicación en Microsoft Entra (una sola vez)
1. portal.azure.com → Microsoft Entra ID → Registros de aplicaciones → Nuevo registro.
2. Nombre: `Dashboard Inspección Contenedores`. Tipos de cuenta: "Solo esta organización".
3. URI de redirección: plataforma **Aplicación de página única (SPA)** con `https://TU-USUARIO.github.io/NOMBRE-REPO/` (con la barra final).
4. Permisos de API → Microsoft Graph → Permisos delegados → `Files.Read.All` → "Conceder consentimiento de administrador" (lo hace TI si no tienes el rol).
5. Copia el **Id de aplicación (cliente)** y pégalo en `config.js` (`clientId`). Si quieres limitarlo a tu tenant, pon su Id en `tenant`.

(Puedes reutilizar el registro del Portal Forecast KAM: solo agrega la URI de redirección nueva y el permiso `Files.Read.All`.)

## 2. Publicar en GitHub
1. Crea el repositorio y sube los 3 archivos a la raíz.
2. Settings → Pages → Source: "Deploy from a branch" → `main` / `(root)`.
3. Abre `https://TU-USUARIO.github.io/NOMBRE-REPO/` e inicia sesión con tu cuenta de INBRISA.

## Notas
- Cada persona solo ve los datos si tiene acceso a la carpeta Matriz compras en SharePoint.
- La página descarga el libro original de la matriz solo para leerlo (no abre sesión de edición, no bloquea a don Luis y no lo modifica).
- Si cambias de carpeta o archivo, actualiza los Id en `config.js`.

- La página pide una contraseña de acceso antes de cargar (se guarda solo su hash en `shim.js`). Es una barrera ligera, porque el código es público; lo que realmente protege los datos es el inicio de sesión Microsoft y los permisos de SharePoint. Para cambiarla, reemplaza `PASS_HASH` por el SHA-256 de la nueva contraseña.
