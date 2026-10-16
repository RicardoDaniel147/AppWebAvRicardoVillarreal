<%
' =====================================================================
'  PoliMusic - Configuracion general de la aplicacion
'  Se incluye SIEMPRE como primera include de cada pagina.
' =====================================================================

Response.Buffer = True
Response.Charset = "utf-8"
Session.CodePage = 65001
Response.CodePage = 65001

' ---------------------------------------------------------------- Identidad
Const SITE_NAME    = "PoliMusic"
Const SITE_TAGLINE = "Streaming"

' ------------------------------------------------------------- Base de datos
Const DB_SERVER = "localhost"
Const DB_NAME   = "BDD_PoliMusic"
Const DB_USER   = "usr_polimusic"
Const DB_PASS   = "Politecnica1"

' ------------------------------------------- Carpeta fisica de los recursos
Const MEDIA_ROOT    = "SongFilesAndImages_resources"
Const DIR_SONGS     = "SongFiles"
Const DIR_USERS     = "UserImages"
Const DIR_PLAYLISTS = "PlaylistImages"
Const DIR_ALBUMS    = "AlbumImages"
Const DIR_ARTISTS   = "ArtistImages"

Const IMG_USER_DEFAULT     = "default.jpg"
Const IMG_PLAYLIST_DEFAULT = "default.jpg"

' ------------------------------------------------------------------- Roles
Const USER_TYPE_NORMAL = 0
Const USER_TYPE_ADMIN  = 1

Const PLAYLIST_PUBLIC  = 0
Const PLAYLIST_PRIVATE = 1

' ------------------------------------------------- Variables globales de pagina
' Las paginas asignan gPage / gTitle antes de incluir header.asp
Dim gPage, gTitle, gPlaylistId
gPage  = ""
gTitle = SITE_NAME
gPlaylistId = 0

%>
