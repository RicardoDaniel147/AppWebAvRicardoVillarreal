# PoliMusic — Aplicación web de streaming musical en **ASP clásico**

Aplicación tipo Spotify construida en **Classic ASP (VBScript)** sobre **SQL Server**,
con estética futurista en negros y azules, dos roles de usuario y despliegue en **IIS**.

---

## 1. Resumen

| Elemento | Detalle |
|---|---|
| Tecnología | Classic ASP (VBScript) — **no requiere .NET** |
| Acceso a datos | ADO / OLE DB (`ADODB.Connection`) con detección automática de proveedor |
| Base de datos | `BDD_PoliMusic` (SQL Server, `localhost`) |
| Sesión | `Session` nativa de ASP |
| Autorización | Variables de sesión + `RequireAdmin()` |
| Subida de archivos | Parser `multipart/form-data` **escrito en VBScript puro** (sin componentes) |
| Front-end | CSS propio (sin CDN), tema oscuro futurista, reproductor global en JS puro |
| Estructura | Plana, con carpetas `admin/`, `assets/`, `includes/` |

---

## 2. Estructura de archivos

```
webapppolimusicv1_asp/
├─ admin/                         # Panel de administración (solo rol ADMIN)
│  ├─ _tabs.asp                   # Pestañas del panel
│  ├─ default.asp                 # Redirige al panel
│  ├─ dashboard.asp               # Métricas del sistema
│  ├─ users.asp        user_edit.asp
│  ├─ songs.asp        song_edit.asp
│  ├─ albums.asp       album_edit.asp
│  ├─ artists.asp      artist_edit.asp
│  ├─ playlists.asp    playlist_edit.asp
│  ├─ catalog.asp      catalog_edit.asp
│  └─ delete.asp                  # Borrado genérico (con borrado en cascada)
├─ assets/
│  ├─ style.css                   # Tema futurista completo
│  ├─ player.js                   # Reproductor global
│  └─ placeholder.asp             # Genera imágenes SVG de relleno
├─ includes/
│  ├─ common.asp                  # Incluye todos los anteriores en orden
│  ├─ config.asp                  # Constantes y configuración
│  ├─ conn.asp                    # Conexión ADO + consultas de ayuda
│  ├─ functions.asp               # Utilidades, escapado SQL, rutas de medios
│  ├─ upload.asp                  # Parser multipart + guardado de archivos
│  ├─ auth.asp                    # Sesión, permisos y protección CSRF
│  ├─ schema.asp                  # Parche de esquema automático
│  ├─ header.asp  footer.asp      # Maqueta general (sidebar + reproductor)
│  ├─ authheader.asp  authfooter.asp  # Maqueta de login/registro
│  └─ songlist.asp                # Listado de canciones reutilizable
├─ SongFilesAndImages_resources/  # Recursos multimedia
│  ├─ SongFiles/ UserImages/ PlaylistImages/ AlbumImages/ ArtistImages/ uploads/
├─ album.asp      albums.asp
├─ artist.asp     artists.asp
├─ default.asp    error.asp      home.asp
├─ login.asp      logout.asp     register.asp
├─ play.asp                      # Contador de reproducciones / redirección al audio
├─ playlist.asp   playlist_edit.asp  playlist_item.asp  playlists.asp
├─ profile.asp    search.asp     songs.asp
├─ BDD_PoliMusic.sql
├─ deploy-iis.ps1                 # Despliegue en IIS (ejecutar como Administrador)
└─ README.md
```

---

## 3. Revisión de la base de datos `BDD_PoliMusic`

9 tablas: `TBL_USER`, `TBL_SONG`, `TBL_ARTIST`, `TBL_ALBUM`, `TBL_TRACKLIST`,
`TBL_PLAYER`, `TBL_PLAYLIST`, `TBL_PLAYLIST_SONG`, `TBL_CATALOG`.

* `TBL_USER.USER_TYPE` → **0 = usuario**, **1 = administrador**
* `TBL_PLAYLIST.TYPE` → **0 = pública**, **1 = privada**
* `TBL_TRACKLIST` (álbum↔canción), `TBL_PLAYER` (canción↔artista) y
  `TBL_PLAYLIST_SONG` (playlist↔canción) son relaciones **N:M**

### Hallazgo: no existía campo de género

El requisito "filtrar por género" no se podía cumplir con el esquema original.
Se agregó de forma **aditiva y no destructiva**:

* `GENRE varchar(50) NULL` en `TBL_SONG` y `TBL_ALBUM`
* Índices `IX_TBL_SONG_GENRE` y `IX_TBL_ALBUM_GENRE`
* Valores iniciales deducidos de los datos existentes

> **Ya está aplicado**, y además la aplicación lo reaplica sola al arrancar
> (`includes/schema.asp`). Si vuelves a ejecutar `BDD_PoliMusic.sql` no se pierde nada.

### Contraseñas

El script original mezcla **texto plano** (`12345678`) y **SHA1 hexadecimal**
(`7C222FB2927D828AF22F592134E8932480637C0D`). El login acepta **ambos formatos**,
y las contraseñas nuevas se guardan siempre como SHA1 (40 caracteres, el máximo de `varchar(50)`).

El hash se calcula **en SQL Server** con `HASHBYTES`, evitando implementar SHA1 en VBScript:

```sql
UPPER(PASSWORD) = UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', 'laClave'), 2))
```

> Detalle importante: `HASHBYTES` distingue `varchar` de `nvarchar`. Como los hashes
> guardados se calcularon sobre `varchar`, el literal debe ir **sin** el prefijo `N`
> (para eso existe la función `SqlAscii()`).

---

## 4. Roles y permisos

| Funcionalidad | Usuario normal | Administrador |
|---|:---:|:---:|
| Iniciar sesión / registrarse | ✅ | ✅ |
| Pantalla principal, explorar, buscar y filtrar | ✅ | ✅ |
| Reproducir (y contar reproducciones) | ✅ | ✅ |
| Álbumes, artistas y sus detalles | ✅ | ✅ |
| Crear y editar **sus propias** playlists | ✅ | ✅ |
| CRUD de **Usuarios** | ❌ | ✅ |
| CRUD de **Canciones** (con subida de audio) | ❌ | ✅ |
| CRUD de **Artistas** (con imagen) | ❌ | ✅ |
| CRUD de **Álbumes** + tracklist | ❌ | ✅ |
| CRUD de **Playlists de cualquier usuario** | ❌ | ✅ |
| CRUD del **Catálogo** | ❌ | ✅ |
| Panel con métricas | ❌ | ✅ |

Un usuario normal que entre a `/admin/...` es redirigido a **Acceso denegado**.
Verificado con pruebas contra la aplicación en ejecución.

---

## 5. Ejecutar en local

Con IIS o IIS Express apuntando a la carpeta del proyecto:

```
http://localhost:<puerto>/default.asp
```

Credenciales:

| Rol | Usuario | Contraseña |
|---|---|---|
| Administrador | `userAdmin` | `12345678` |
| Usuario normal | `userNormal` | `12345678` |

---

## 6. Despliegue en IIS

### 6.1 Activar Classic ASP (una sola vez)

`Panel de control` → `Programas` → `Activar o desactivar características de Windows` →
`Internet Information Services` → `World Wide Web Services` → `Application Development Features` → **ASP**.

Además, para servir CSS/JS/imágenes: `Common HTTP Features` → **Static Content** y **Default Document**.

> Classic ASP **no necesita .NET ni el Hosting Bundle**. Solo `asp.dll`.

### 6.2 Desplegar

Abre **PowerShell como Administrador** y ejecuta:

```powershell
cd C:\Users\Ricardo\WEBREPOS\AppWebAvRicardoVillarreal\webapppolimusicv1_asp
.\deploy-iis.ps1
```

Por defecto despliega en `C:\inetpub\wwwroot\PoliMusic`, crea el grupo de aplicaciones
`PoliMusicAppPool` y publica el sitio en el puerto **8080**.

```powershell
# Publicar directamente en la raíz del sitio, en el puerto 80
.\deploy-iis.ps1 -Target "C:\inetpub\wwwroot" -Port 80
```

El script:

* Verifica privilegios de administrador y que Classic ASP esté instalado.
* **Copia solo los archivos de la aplicación** (excluye `docs/`, `.workbuddy-ai/`, este README, etc.).
* Respalda la carpeta destino si ya existe, en lugar de borrarla.
* Crea el grupo de aplicaciones como **No Managed Code** y el sitio IIS.
* Ajusta permisos NTFS para `IIS_IUSRS` y la identidad del grupo de aplicaciones.
* **Habilita `enableParentPaths`** — imprescindible: las páginas de `/admin` usan
  `#include file="../includes/..."`, y IIS lo bloquea por defecto.
* Sube los límites de subida de ASP e IIS a 100 MB (configurable con `-MaxUploadMb`).
* Agrega `default.asp` a los documentos predeterminados.

### 6.3 Sobre la ruta indicada

Se interpretó `C:\initpub\wwroot` como **`C:\inetpub\wwwroot`**.
Si querías otra ruta, pásala con `-Target`.

---

## 7. Limitaciones conocidas de los recursos

La carpeta `SongFilesAndImages_resources` está **incompleta** respecto a lo que
referencia la base de datos:

| Recurso | Referenciado en BD | Disponible |
|---|---|---|
| Canciones | 5 (`bensound-*.mp3`) | **1** (`bensound-adventure.mp3`) |
| `UserImages` | `default.jpg`, `root.png` | ✅ ambos |
| `PlaylistImages` | `default.jpg`, `playlist1-5.png` | ✅ todos |
| `AlbumImages` | `album1-5.png` | ❌ ninguno |
| `ArtistImages` | `artist1-5` | ❌ ninguno |

La aplicación está preparada para esto:

* **Imágenes faltantes** → `assets/placeholder.asp` genera un **SVG con degradado e
  iniciales**, así nunca se ven imágenes rotas.
* **Audio faltante** → la fila muestra la etiqueta *"Sin archivo"* y al intentar
  reproducir aparece un aviso claro, sin romper la interfaz.
* **Subidas nuevas** → desde el panel de administración puedes subir los MP3 e imágenes
  que faltan; se guardan automáticamente en la carpeta correcta.

---

## 8. Verificación realizada

La aplicación se levantó en IIS Express contra la base de datos real:

* ✅ Arranque correcto y parche de esquema aplicado automáticamente.
* ✅ Redirección de anónimos a `login.asp`.
* ✅ Login de administrador, de usuario normal y de un usuario con **contraseña SHA1**.
* ✅ Registro público con **subida de foto** (multipart propio).
* ✅ 200 en todas las páginas públicas y del panel `/admin`.
* ✅ Bloqueo de `/admin/*` para el usuario normal (redirección a *Acceso denegado*).
* ✅ CRUD completo: crear, editar y eliminar **usuarios, canciones, álbumes, artistas,
  playlists y catálogo**, con borrado en cascada de las relaciones.
* ✅ Protección: el administrador no puede eliminarse a sí mismo ni quitarse el rol.
* ✅ Contador de reproducciones (`POST play.asp`) y validación CSRF (403 sin token).
* ✅ Filtros por género, artista y álbum; pestañas de resultados.
* ✅ Los datos de prueba se eliminaron: la base quedó en su estado original
  (4 usuarios, 5 canciones, 5 artistas, 5 álbumes, 5 playlists, 6 entradas de catálogo).

---

## 9. Seguridad implementada

* **Escapado SQL** en todos los valores de texto (`SqlText`/`SqlAscii`) y conversión
  forzada a entero en los numéricos (`ToInt`/`SqlInt0`).
* **Protección CSRF**: token por sesión, presente en todos los formularios y validado
  en cada POST (`CsrfFieldInline` / `CsrfCheck`). El reproductor lo envía por cabecera.
* **Autorización** por `RequireLogin()` / `RequireAdmin()` en cada página.
* **Comprobación de propiedad**: un usuario solo edita sus propias playlists y su perfil.
* **Validación de subidas**: extensión permitida y nombre de archivo generado
  (nunca se usa el nombre original del cliente).
* **Codificación de salida** con `H()` en todo lo que se imprime (protección XSS).

---

## 10. Notas técnicas para el mantenimiento

* `includes/common.asp` es el único include que necesitan las páginas públicas;
  las de `admin/` usan `../includes/common.asp`.
* `SqlCanciones()`, `SqlAlbumes()`, `SqlArtistas()` y `SqlPlaylists()` (en `functions.asp`)
  centralizan las consultas con sus agregados; las páginas solo añaden el `WHERE`.
* El reproductor funciona sobre atributos `data-*` de cada fila (`.pm-row`), así que
  cualquier listado nuevo solo necesita usar `includes/songlist.asp`.
* Las rutas de la base de datos (`~/songFiles/x.mp3`) se traducen a URL reales con
  `MediaUrl()` / `AudioUrl()`, que además comprueban que el archivo exista.

---

## 11. Solución de problemas

### Error 404.0 al abrir `http://localhost/login.asp`

**Síntoma:** la página de error muestra `Handler: ASPClassic` y
`Physical Path: C:\inetpub\wwwroot\login.asp`, pero el archivo no está ahí.

**Causa:** se copió la **carpeta** del proyecto dentro de `wwwroot` en lugar de su
**contenido**. Los archivos quedan en `C:\inetpub\wwwroot\webapppolimusicv1_asp\...`,
así que `/login.asp` no existe en la raíz del sitio.

**Solución — elige una de las dos:**

**A) Publicar en la raíz del sitio** (así la app abre en `http://localhost/`):

```powershell
# PowerShell como Administrador
cd C:\Users\Ricardo\WEBREPOS\AppWebAvRicardoVillarreal\webapppolimusicv1_asp
.\deploy-iis.ps1 -Target "C:\inetpub\wwwroot" -Port 80 -SkipSite
```

Este script copia solo los archivos de la aplicación (no `docs/`, ni este README,
ni el propio script) y **no borra** el contenido existente de `wwwroot`.

**B) Mantener la subcarpeta** (la app abre en `http://localhost/webapppolimusicv1_asp/`):

Basta con volver a copiar el proyecto a esa carpeta. Es válido, pero recuerda que
la URL siempre incluirá el nombre de la carpeta.

> En **ambos** casos hay que aplicar la configuración de ASP (siguiente punto).

### La página se ve sin estilos y `/admin` devuelve 500

**Causa doble:**

1. Si la carpeta **no es una aplicación IIS**, `APPL_MD_PATH` apunta a la raíz del sitio
   y la aplicación generaba enlaces a `/assets/...` en vez de `/miCarpeta/assets/...`.
   **Ya está corregido**: `AppPath()` (en `includes/functions.asp`) ahora deduce la ruta
   real a partir de la URL del script, por lo que funciona igual en la raíz o en una
   subcarpeta. **Necesitas volver a copiar el proyecto** para llevar esta corrección.
2. Las páginas de `/admin` usan `#include file="../includes/common.asp"`, y **IIS bloquea
   `..` por defecto**. Hay que habilitar las rutas padre:

```powershell
# PowerShell como Administrador
Import-Module WebAdministration
Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
    -Filter "system.webServer/asp" -Name "enableParentPaths" -Value $true

# Subir el limite de subida de archivos a 100 MB (opcional pero recomendado)
Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
    -Filter "system.webServer/asp/limits" -Name "maxRequestEntityAllowed" -Value 104857600
Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
    -Filter "system.webServer/asp/limits" -Name "bufferingLimit" -Value 104857600
Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' `
    -Filter "system.webServer/security/requestFiltering/requestLimits" `
    -Name "maxAllowedContentLength" -Value 104857600
```

O con `appcmd`:

```
%windir%\system32\inetsrv\appcmd set config /section:asp /enableParentPaths:true
```

El script `deploy-iis.ps1` aplica todo esto automáticamente.

### Otros errores frecuentes

| Error | Causa | Solución |
|---|---|---|
| `HTTP 500.19` / handler no encontrado | Classic ASP no está activado | Activar la característica **ASP** de IIS |
| `ASP 0131 Disallowed Parent Path` | Falta `enableParentPaths` | Ver el punto anterior |
| `ASP 0116 Missing close of script delimiter` | Un `%>` suelto o un `%>` dentro de un comentario `'` | Revisar el balance de `<%` / `%>` del archivo que indica el error |
| `Microsoft VBScript runtime error 800a000d` | Tipo de dato inesperado | Suele venir de un valor NULL; revisar la consulta |
| Error al subir archivos grandes | Límite de ASP/IIS | Aplicar los límites de arriba |
| Las imágenes de álbum/artista salen como degradado con iniciales | Faltan esos archivos | Es el comportamiento previsto; súbelos desde el panel |

### Comprobar que Classic ASP responde

Crea un archivo `prueba.asp` en la raíz del sitio con:

```asp
<% Response.Write "ASP funciona: " & Now() %>
```

Si al abrirlo ves la fecha, Classic ASP está bien configurado. **Bórralo después.**
