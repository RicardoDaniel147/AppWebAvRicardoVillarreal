# MEMORY.md — AppWebAvRicardoVillarreal (PoliMusic)

## Proyecto
Aplicación web de streaming musical tipo Spotify ("PoliMusic"), proyecto de Web Avanzada.
Raíz: `C:\Users\Ricardo\WEBREPOS\AppWebAvRicardoVillarreal`.
**Versión actual: ASP clásico (VBScript) en `webapppolimusicv1_asp/`.**
(La versión anterior en ASP.NET Core fue eliminada por el usuario el 2026-10-15.)

## Convenciones y decisiones fijadas

- **Tecnología**: **Classic ASP (VBScript)**, NO ASP.NET. No requiere .NET ni Hosting Bundle;
  solo la característica "ASP" de IIS (`asp.dll`).
- **Estructura plana** (la pidió el usuario así, según captura): páginas `.asp` en la raíz
  + carpetas `admin/`, `assets/`, `includes/`, `SongFilesAndImages_resources/`.
- **Acceso a datos**: ADO / OLE DB con detección automática de proveedor
  (`includes/conn.asp` prueba SQLOLEDB → MSOLEDBSQL → MSDASQL+ODBC17/18 y cachea el que funcione).
  El usuario `usr_polimusic` / `Politecnica1` funciona con todos.
- **Estilos**: CSS propio en `assets/style.css` (tema oscuro negro/azul futurista).
  **Prohibido usar CDN** — todo local. Fuente del sistema.
- **Idioma de la interfaz**: español, **sin tildes** en el código/HTML por compatibilidad
  de codificación (el usuario escribe sin tildes).
- **Base de datos**: `BDD_PoliMusic` en `localhost`. Constantes en `includes/config.asp`.
- **Roles**: `TBL_USER.USER_TYPE` → 0 = usuario normal, 1 = administrador.
  `RequireAdmin()` protege todo `/admin`.
- **Contraseñas**: el login acepta texto plano y SHA1 hex. **El hash se calcula en SQL Server**
  con `HASHBYTES` (no hay SHA1 en VBScript). OJO: `HASHBYTES` distingue varchar de nvarchar,
  por eso el literal va sin prefijo `N` → usar la función `SqlAscii()`.
- **Esquema**: `includes/schema.asp` aplica parches idempotentes al arrancar
  (columnas `GENRE`, índices, usuario admin por defecto). Si se re-ejecuta
  `BDD_PoliMusic.sql` no se pierde nada.
- **Imágenes faltantes**: usar siempre `MediaUrl()` en las vistas; devuelve la imagen real
  o `assets/placeholder.asp` (SVG generado). Nunca poner la ruta cruda de la BD.
- **Subidas**: `includes/upload.asp` tiene un parser multipart **en VBScript puro**
  (sin Persits ni componentes). Guardar con `SaveUpload(campo, carpeta, "|.jpg|.png|")`.
  Las páginas que reciben POST multipart deben llamar `ParseMultipart` **antes** de `CsrfCheck`.
- **CSRF**: token por sesión. Formularios con `CsrfFieldInline()`, validación con `CsrfCheck()`.
- **Includes**: las páginas de la raíz usan `<!--#include file="includes/common.asp" -->`;
  las de `admin/` usan `<!--#include file="../includes/common.asp" -->`.
  **IIS bloquea `..` por defecto** → hay que activar `enableParentPaths=true` (el script
  de despliegue lo hace; en IIS Express hay que ponerlo a mano en el config).

## Despliegue
- Destino: IIS. El usuario escribió "C:\initpub\wwroot" → interpretado como `C:\inetpub\wwwroot`.
- Script: `deploy-iis.ps1` **como Administrador**. Copia solo los archivos de la app,
  crea app pool "No Managed Code", sitio, permisos, `enableParentPaths`, límites de subida
  de 100 MB y `default.asp` como documento predeterminado.
- El usuario **no tiene privilegios de administrador** en su sesión normal.

## Credenciales de prueba
- Administrador: `userAdmin` / `12345678`
- Usuario normal: `userNormal` / `12345678`
- Usuario con hash SHA1: `username1` / `12345678`

## Trampas encontradas en Classic ASP (útiles para futuras ediciones)
1. **`%>` dentro de un comentario** `'` cierra el bloque de script antes de tiempo.
   Verificar siempre el balance de `<%` / `%>`.
2. **`IsObject(Nothing)` devuelve True** en VBScript → usar `x Is Nothing`.
3. Los includes puros de código deben abrir con `<%` **y cerrar con `%>`**.
4. `#include file="../..."` requiere `enableParentPaths=true` en IIS.
5. VBScript **no tiene `IIf`** → está definida en `includes/functions.asp`.
6. `Response.Redirect` debe ejecutarse **antes** de incluir `header.asp`.
