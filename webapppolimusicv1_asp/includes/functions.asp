<%
' =====================================================================
'  PoliMusic - Funciones auxiliares compartidas
' =====================================================================

Dim gFsoObj
Set gFsoObj = Nothing

' Objeto FileSystemObject reutilizado durante la peticion.
Function GetFSO()
    If gFsoObj Is Nothing Then
        Set gFsoObj = Server.CreateObject("Scripting.FileSystemObject")
    End If
    Set GetFSO = gFsoObj
End Function

' --------------------------------------------------------------- salida HTML
' Codifica una cadena para insertarla con seguridad en el HTML.
Function H(v)
    Dim s
    s = "" & v
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

' Codifica una cadena para usarla dentro de un atributo o de JavaScript.
Function JS(v)
    Dim s
    s = "" & v
    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), " ")
    s = Replace(s, Chr(10), " ")
    JS = s
End Function

' ------------------------------------------------------------------- numeros
' Conversion segura a entero con valor por defecto.
Function ToInt(v, defVal)
    Dim n
    n = defVal
    On Error Resume Next
    If Not IsNull(v) Then
        If Trim("" & v) <> "" Then n = CLng(v)
    End If
    If Err.Number <> 0 Then n = defVal
    Err.Clear
    On Error GoTo 0
    ToInt = n
End Function

' Formatea un numero con separador de miles.
Function Num(v)
    Dim n
    n = ToInt(v, 0)
    Num = FormatNumber(n, 0, -1, 0, -1)
End Function

' ------------------------------------------------------------------- fechas
Function FechaCorta(v)
    If IsNull(v) Or "" & v = "" Then
        FechaCorta = "-"
    Else
        On Error Resume Next
        FechaCorta = Right("0" & Day(v), 2) & "/" & Right("0" & Month(v), 2) & "/" & Year(v)
        If Err.Number <> 0 Then FechaCorta = "-"
        Err.Clear
        On Error GoTo 0
    End If
End Function

' Devuelve el valor para un input type="date".
Function FechaInput(v)
    If IsNull(v) Or "" & v = "" Then
        FechaInput = ""
    Else
        On Error Resume Next
        FechaInput = Year(v) & "-" & Right("0" & Month(v), 2) & "-" & Right("0" & Day(v), 2)
        If Err.Number <> 0 Then FechaInput = ""
        Err.Clear
        On Error GoTo 0
    End If
End Function

' --------------------------------------------------------------- validaciones
Function IsValidEmail(s)
    Dim re
    IsValidEmail = False
    s = Trim("" & s)
    If Len(s) < 5 Or InStr(s, "@") = 0 Then Exit Function
    On Error Resume Next
    Set re = New RegExp
    re.Pattern = "^[A-Za-z0-9._%\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    re.IgnoreCase = True
    IsValidEmail = re.Test(s)
    Set re = Nothing
    Err.Clear
    On Error GoTo 0
End Function

' ------------------------------------------------------------------- lectura
' Lee un valor de Request (GET o POST) y lo recorta.
Function R(v)
    R = Trim("" & Request(v))
End Function

' Lee del formulario.
Function Frm(v)
    Frm = Trim("" & Request.Form(v))
End Function

' Lee de la cadena de consulta.
Function Qry(v)
    Qry = Trim("" & Request.QueryString(v))
End Function

' --------------------------------------------------------- SQL seguro (escape)
' Devuelve un literal de texto SQL entre comillas simples, escapado.
Function SqlText(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        SqlText = "NULL"
        Exit Function
    End If
    s = "" & v
    s = Replace(s, Chr(0), "")
    s = Replace(s, "'", "''")
    SqlText = "N'" & s & "'"
End Function

' Igual que SqlText pero devuelve NULL cuando la cadena esta vacia.
Function SqlTextOrNull(v)
    If IsNull(v) Or IsEmpty(v) Or Trim("" & v) = "" Then
        SqlTextOrNull = "NULL"
    Else
        SqlTextOrNull = SqlText(v)
    End If
End Function

' Literal de texto SIN el prefijo N (varchar).
' Es imprescindible para HASHBYTES: hashear un nvarchar da un resultado distinto
' al de un varchar, y la base guarda los hashes calculados sobre varchar.
Function SqlAscii(v)
    Dim s
    s = "" & v
    s = Replace(s, Chr(0), "")
    s = Replace(s, "'", "''")
    SqlAscii = "'" & s & "'"
End Function

' Devuelve un literal entero, o NULL si no hay valor.
Function SqlInt(v)
    If IsNull(v) Or IsEmpty(v) Or Trim("" & v) = "" Then
        SqlInt = "NULL"
    Else
        SqlInt = CStr(ToInt(v, 0))
    End If
End Function

' Devuelve un literal entero, o 0 si no hay valor.
Function SqlInt0(v)
    SqlInt0 = CStr(ToInt(v, 0))
End Function

' Devuelve un literal de fecha, o NULL.
Function SqlDate(v)
    If IsNull(v) Or IsEmpty(v) Or Trim("" & v) = "" Then
        SqlDate = "NULL"
    Else
        On Error Resume Next
        Dim d
        d = CDate(v)
        If Err.Number <> 0 Then
            SqlDate = "NULL"
        Else
            SqlDate = "'" & Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "'"
        End If
        Err.Clear
        On Error GoTo 0
    End If
End Function

' ------------------------------------------------------------- rutas / enlaces
' Ruta virtual de la aplicacion. Funciona tanto si la app esta en la raiz del
' sitio como si esta dentro de una subcarpeta, y no necesita que la carpeta sea
' una "aplicacion IIS".
Function AppPath()
    Dim md, idx, sn, p

    ' 1) Metodo preferido: APPL_MD_PATH (correcto si la carpeta ES una aplicacion IIS).
    md = "" & Request.ServerVariables("APPL_MD_PATH")
    idx = InStrRev(md, "/ROOT")
    If idx > 0 Then
        p = Mid(md, idx + 5)
    Else
        p = ""
    End If
    If p = "/" Then p = ""

    ' 2) Si APPL_MD_PATH apunta a la raiz del sitio pero las paginas viven en una
    '    subcarpeta (caso tipico al copiar la carpeta dentro de wwwroot), se deduce
    '    la raiz a partir de la ruta virtual del script que se esta ejecutando.
    If p = "" Then
        sn = "" & Request.ServerVariables("SCRIPT_NAME")   ' p.ej. /miapp/login.asp
        If sn <> "" Then
            idx = InStr(LCase(sn), "/admin/")
            If idx > 0 Then
                ' Las paginas de /admin estan un nivel mas abajo
                p = Left(sn, idx - 1)
            Else
                idx = InStrRev(sn, "/")
                If idx > 1 Then p = Left(sn, idx - 1)
            End If
            If p = "/" Then p = ""
        End If
    End If

    AppPath = p
End Function

' Construye una URL absoluta dentro de la aplicacion (rel sin barra inicial).
Function AppUrl(rel)
    AppUrl = AppPath() & "/" & rel
End Function

' URL de un recurso de assets.
Function Asset(rel)
    Asset = AppPath() & "/assets/" & rel
End Function

' ------------------------------------------------------------- multimedia
' Traduce el nombre de carpeta que usa la base de datos al nombre fisico real.
Function FolderFor(dbFolder)
    Select Case LCase("" & dbFolder)
        Case "songfiles"      : FolderFor = DIR_SONGS
        Case "userimages"     : FolderFor = DIR_USERS
        Case "playlistimages" : FolderFor = DIR_PLAYLISTS
        Case "albumimages"    : FolderFor = DIR_ALBUMS
        Case "artistimages"   : FolderFor = DIR_ARTISTS
        Case Else             : FolderFor = ""
    End Select
End Function

' Ruta fisica de la carpeta de recursos.
Function MediaPhysicalRoot()
    MediaPhysicalRoot = Server.MapPath(AppPath() & "/" & MEDIA_ROOT)
End Function

Function FileExists(path)
    On Error Resume Next
    FileExists = GetFSO().FileExists(path)
    Err.Clear
    On Error GoTo 0
End Function

' URL del generador de imagenes de relleno.
Function PlaceholderUrl(kind, seed)
    PlaceholderUrl = AppPath() & "/assets/placeholder.asp?kind=" & Server.URLEncode("" & kind) & _
                     "&seed=" & Server.URLEncode("" & seed)
End Function

' Convierte una ruta de la base de datos (~/songFiles/x.mp3) en una URL real.
' Si el archivo no existe devuelve la imagen de relleno.
Function MediaUrl(dbPath, kind, seed)
    Dim p, parts, folder, file, physical
    p = Trim("" & dbPath)
    p = Replace(p, "\", "/")
    If Left(p, 2) = "~/" Then p = Mid(p, 3)
    If Left(p, 1) = "/" Then p = Mid(p, 2)

    If p = "" Then
        MediaUrl = PlaceholderUrl(kind, seed)
        Exit Function
    End If

    parts = Split(p, "/")
    If UBound(parts) >= 1 Then
        folder = FolderFor(parts(0))
        file = parts(UBound(parts))
    Else
        folder = ""
        file = parts(0)
    End If

    If folder = "" Or file = "" Then
        MediaUrl = PlaceholderUrl(kind, seed)
        Exit Function
    End If

    physical = MediaPhysicalRoot() & "\" & folder & "\" & file
    If FileExists(physical) Then
        MediaUrl = AppPath() & "/" & MEDIA_ROOT & "/" & folder & "/" & file
    Else
        MediaUrl = PlaceholderUrl(kind, seed)
    End If
End Function

' URL del archivo de audio, o cadena vacia si no existe fisicamente.
Function AudioUrl(dbPath)
    Dim p, parts, folder, file, physical
    AudioUrl = ""
    p = Trim("" & dbPath)
    p = Replace(p, "\", "/")
    If Left(p, 2) = "~/" Then p = Mid(p, 3)
    If Left(p, 1) = "/" Then p = Mid(p, 2)
    If p = "" Then Exit Function

    parts = Split(p, "/")
    If UBound(parts) >= 1 Then
        folder = FolderFor(parts(0))
        file = parts(UBound(parts))
    Else
        folder = DIR_SONGS
        file = parts(0)
    End If
    If folder = "" Or file = "" Then Exit Function

    physical = MediaPhysicalRoot() & "\" & folder & "\" & file
    If FileExists(physical) Then
        AudioUrl = AppPath() & "/" & MEDIA_ROOT & "/" & folder & "/" & file
    End If
End Function

Function AudioExists(dbPath)
    AudioExists = (Len(AudioUrl(dbPath)) > 0)
End Function

' --------------------------------------------------------------- textos cortos
Function Iniciales(texto, maxLen)
    Dim s, i, c, out
    s = Trim("" & texto)
    If s = "" Then
        Iniciales = "?"
        Exit Function
    End If
    out = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c <> " " Then
            out = out & UCase(c)
            If Len(out) >= maxLen Then Exit For
        End If
    Next
    Iniciales = out
End Function

Function Recortar(texto, maxLen)
    Dim s
    s = Trim("" & texto)
    If Len(s) <= maxLen Then
        Recortar = s
    Else
        Recortar = Left(s, maxLen - 1) & "..."
    End If
End Function

' Nombre para mostrar de un artista a partir de nombre y apellido.
Function NombreArtista(nombre, apellido)
    Dim s
    s = Trim("" & nombre)
    If Trim("" & apellido) <> "" Then s = s & " " & Trim("" & apellido)
    NombreArtista = s
End Function

' Etiqueta del tipo de usuario.
Function TipoUsuarioNombre(t)
    If ToInt(t, 0) = USER_TYPE_ADMIN Then
        TipoUsuarioNombre = "Administrador"
    Else
        TipoUsuarioNombre = "Usuario"
    End If
End Function

' Etiqueta del tipo de playlist.
Function TipoPlaylistNombre(t)
    If ToInt(t, 0) = PLAYLIST_PRIVATE Then
        TipoPlaylistNombre = "Privada"
    Else
        TipoPlaylistNombre = "Publica"
    End If
End Function

' ------------------------------------------------------- genero de las canciones
' Lista de generos distintos de las canciones (array de cadenas).
Function GenerosCanciones()
    Dim rs, s
    s = ""
    Set rs = Query("SELECT DISTINCT GENRE FROM TBL_SONG " & _
                   "WHERE GENRE IS NOT NULL AND LTRIM(RTRIM(GENRE)) <> '' ORDER BY GENRE")
    Do While Not rs.EOF
        If s <> "" Then s = s & "|"
        s = s & ("" & rs("GENRE"))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If s = "" Then
        GenerosCanciones = Array()
    Else
        GenerosCanciones = Split(s, "|")
    End If
End Function

' Lista de generos distintos de los albumes.
Function GenerosAlbumes()
    Dim rs, s
    s = ""
    Set rs = Query("SELECT DISTINCT GENRE FROM TBL_ALBUM " & _
                   "WHERE GENRE IS NOT NULL AND LTRIM(RTRIM(GENRE)) <> '' ORDER BY GENRE")
    Do While Not rs.EOF
        If s <> "" Then s = s & "|"
        s = s & ("" & rs("GENRE"))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    If s = "" Then
        GenerosAlbumes = Array()
    Else
        GenerosAlbumes = Split(s, "|")
    End If
End Function

' Devuelve True si el array contiene el valor (comparacion sin distinguir mayusculas).
Function EnLista(lista, valor)
    Dim i
    EnLista = False
    If Not IsArray(lista) Then Exit Function
    For i = 0 To UBound(lista)
        If LCase("" & lista(i)) = LCase("" & valor) Then
            EnLista = True
            Exit Function
        End If
    Next
End Function

' ------------------------------------------------------------- utilidades HTML
' Mensaje flash guardado en sesion.
Sub Flash(msg, tipo)
    Session("flash_msg") = "" & msg
    Session("flash_tipo") = "" & tipo
End Sub

Function TakeFlash()
    Dim m
    m = "" & Session("flash_msg")
    Session("flash_msg") = ""
    TakeFlash = m
End Function

Function TakeFlashTipo()
    Dim t
    t = "" & Session("flash_tipo")
    Session("flash_tipo") = ""
    If t = "" Then t = "ok"
    TakeFlashTipo = t
End Function

' ------------------------------------------------------------ utilidades varias
' Equivalente a IIf (VBScript no lo trae de serie).
Function IIf(cond, siVerdadero, siFalso)
    If cond Then
        IIf = siVerdadero
    Else
        IIf = siFalso
    End If
End Function

' Marca el enlace activo del menu lateral.
Function NavActivo(nombre)
    If LCase("" & gPage) = LCase("" & nombre) Then
        NavActivo = "active"
    Else
        NavActivo = ""
    End If
End Function

' Clase CSS del badge segun el tipo de usuario.
Function BadgeUsuario(esAdmin)
    If esAdmin Then
        BadgeUsuario = "pm-badge pm-badge-admin"
    Else
        BadgeUsuario = "pm-badge pm-badge-user"
    End If
End Function

' Clase CSS del badge segun el tipo de playlist.
Function BadgePlaylist(esPublica)
    If ToInt(esPublica, 0) = PLAYLIST_PUBLIC Then
        BadgePlaylist = "pm-badge pm-badge-public"
    Else
        BadgePlaylist = "pm-badge pm-badge-private"
    End If
End Function

' Valor para un checkbox (1 / 0).
Function Checked(v)
    If ToInt(v, 0) <> 0 Then
        Checked = "checked=""checked"""
    Else
        Checked = ""
    End If
End Function

' Opcion de un select, marcando la seleccionada.
Function OptionSel(valor, etiqueta, actual)
    Dim sel
    If LCase("" & valor) = LCase("" & actual) Then
        sel = " selected=""selected"""
    Else
        sel = ""
    End If
    OptionSel = "<option value=""" & H(valor) & """" & sel & ">" & H(etiqueta) & "</option>"
End Function

' Marca "selected" si coinciden.
Function SelIf(valor, actual)
    If LCase("" & valor) = LCase("" & actual) Then
        SelIf = " selected=""selected"""
    Else
        SelIf = ""
    End If
End Function

' ------------------------------------------------- consultas reutilizables
' Devuelve el SQL de canciones incluyendo los artistas asociados y una imagen.
Function SqlCanciones(condicion, orden, topN)
    Dim sql, top
    top = ""
    If ToInt(topN, 0) > 0 Then top = "TOP " & ToInt(topN, 0) & " "

    sql = "SELECT " & top & _
          "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, s.PLAYS, s.GENRE, " & _
          "ISNULL(STUFF((SELECT ', ' + a.ARTIST_NAME + ISNULL(' ' + a.ARTIST_LASTNAME,'') " & _
          "FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
          "WHERE p.ID_SONG = s.ID_SONG FOR XML PATH('')),1,2,''),'Artista desconocido') AS ARTISTAS, " & _
          "(SELECT TOP 1 a2.ARTIST_IMAGE FROM TBL_PLAYER p2 " & _
          "INNER JOIN TBL_ARTIST a2 ON a2.ID_ARTIST = p2.ID_ARTIST " & _
          "WHERE p2.ID_SONG = s.ID_SONG ORDER BY a2.ARTIST_NAME) AS ARTIST_IMAGE " & _
          "FROM TBL_SONG s "

    If condicion <> "" Then sql = sql & "WHERE " & condicion & " "
    If orden <> "" Then sql = sql & "ORDER BY " & orden & " "
    SqlCanciones = sql
End Function

' Devuelve el SQL de albumes con el numero de canciones.
Function SqlAlbumes(condicion, orden, topN)
    Dim sql, top
    top = ""
    If ToInt(topN, 0) > 0 Then top = "TOP " & ToInt(topN, 0) & " "

    sql = "SELECT " & top & _
          "a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, a.GENRE, " & _
          "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS CANCIONES " & _
          "FROM TBL_ALBUM a "

    If condicion <> "" Then sql = sql & "WHERE " & condicion & " "
    If orden <> "" Then sql = sql & "ORDER BY " & orden & " "
    SqlAlbumes = sql
End Function

' Devuelve el SQL de artistas con el numero de canciones.
Function SqlArtistas(condicion, orden, topN)
    Dim sql, top
    top = ""
    If ToInt(topN, 0) > 0 Then top = "TOP " & ToInt(topN, 0) & " "

    sql = "SELECT " & top & _
          "a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
          "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = a.ID_ARTIST) AS CANCIONES " & _
          "FROM TBL_ARTIST a "

    If condicion <> "" Then sql = sql & "WHERE " & condicion & " "
    If orden <> "" Then sql = sql & "ORDER BY " & orden & " "
    SqlArtistas = sql
End Function

' Devuelve el SQL de playlists con propietario y numero de canciones.
Function SqlPlaylists(condicion, orden, topN)
    Dim sql, top
    top = ""
    If ToInt(topN, 0) > 0 Then top = "TOP " & ToInt(topN, 0) & " "

    sql = "SELECT " & top & _
          "pl.ID_PLAYLIST, pl.ID_USER, pl.TITLE, pl.CREATION_DATE, pl.TYPE, pl.PHOTO, " & _
          "u.USERNAME AS PROPIETARIO, " & _
          "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS CANCIONES " & _
          "FROM TBL_PLAYLIST pl LEFT JOIN TBL_USER u ON u.ID_USER = pl.ID_USER "

    If condicion <> "" Then sql = sql & "WHERE " & condicion & " "
    If orden <> "" Then sql = sql & "ORDER BY " & orden & " "
    SqlPlaylists = sql
End Function

%>
