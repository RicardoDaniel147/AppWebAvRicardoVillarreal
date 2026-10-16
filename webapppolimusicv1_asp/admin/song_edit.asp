<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar cancion (solo administradores)
' =====================================================================
RequireAdmin

Dim seId, seNew, seName, sePath, sePlays, seGenre, seArtists, seErr
Dim rsS, rsArtistas, seUp, seArr, sei, seIds

seId    = ToInt(Qry("id"), 0)
seNew   = (seId = 0)
seErr   = ""
seName  = ""
sePath  = ""
sePlays = 0
seGenre = ""
seIds   = ""

If Not seNew Then
    Set rsS = Query("SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS, GENRE FROM TBL_SONG WHERE ID_SONG = " & seId)
    If rsS.EOF Then
        rsS.Close : Set rsS = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    seName  = "" & rsS("SONG_NAME")
    sePath  = "" & rsS("SONG_PATH")
    sePlays = ToInt(rsS("PLAYS"), 0)
    seGenre = "" & rsS("GENRE")
    rsS.Close : Set rsS = Nothing

    Set rsS = Query("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = " & seId)
    Do While Not rsS.EOF
        seIds = seIds & IIf(seIds = "", "", ",") & rsS("ID_ARTIST")
        rsS.MoveNext
    Loop
    rsS.Close : Set rsS = Nothing
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    seName  = PostVal("songname")
    sePath  = PostVal("songpath")
    sePlays = ToInt(PostVal("plays"), 0)
    seGenre = PostVal("genre")
    seIds   = PostVal("artists")

    If Len(Trim(seName)) = 0 Then
        seErr = "El nombre de la cancion es obligatorio."
    ElseIf sePlays < 0 Then
        seErr = "Las reproducciones no pueden ser negativas."
    End If

    If seErr = "" Then
        If HasUpload("audio") Then
            seUp = SaveUpload("audio", DIR_SONGS, "|.mp3|.wav|.ogg|.m4a|.aac|.flac|")
            If seUp <> "" Then sePath = seUp
        End If

        If Trim(sePath) = "" Then
            seErr = "Sube un archivo de audio o indica la ruta del archivo."
        End If
    End If

    If seErr = "" Then
        If seNew Then
            seId = ExecuteInsert("INSERT INTO TBL_SONG(SONG_NAME, SONG_PATH, PLAYS, GENRE) VALUES(" & _
                   SqlText(Trim(seName)) & ", " & SqlText(Trim(sePath)) & ", " & SqlInt0(sePlays) & ", " & SqlTextOrNull(seGenre) & ")")
            Flash "Cancion """ & Trim(seName) & """ creada correctamente.", "ok"
        Else
            Execute "UPDATE TBL_SONG SET SONG_NAME = " & SqlText(Trim(seName)) & _
                    ", SONG_PATH = " & SqlText(Trim(sePath)) & _
                    ", PLAYS = " & SqlInt0(sePlays) & _
                    ", GENRE = " & SqlTextOrNull(seGenre) & _
                    " WHERE ID_SONG = " & seId
            Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & seId
            Flash "Cancion """ & Trim(seName) & """ actualizada.", "ok"
        End If

        If seIds <> "" Then
            seArr = Split(seIds, ",")
            For sei = 0 To UBound(seArr)
                If ToInt(seArr(sei), 0) > 0 Then
                    Execute "INSERT INTO TBL_PLAYER(ID_SONG, ID_ARTIST) VALUES(" & seId & ", " & ToInt(seArr(sei), 0) & ")"
                End If
            Next
        End If

        Response.Redirect AppUrl("admin/songs.asp")
        Response.End
    End If
End If

Set rsArtistas = Query(SqlArtistas("", "a.ARTIST_NAME", 0))

gPage = "admin-songs"
gTitle = IIf(seNew, "Nueva cancion", "Editar cancion")
%>
<!--#include file="../includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/songs.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a canciones
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub"><%=IIf(seNew, "Registra una nueva pista subiendo el archivo o indicando su ruta.", "Actualiza los datos de la pista, su genero y sus artistas.")%></p>

<% If seErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(seErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("admin/song_edit.asp")%>?id=<%=seId%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Datos de la pista</h3>

        <div class="pm-form-grid">
            <div class="pm-field">
                <label for="songname">Nombre de la cancion</label>
                <input type="text" id="songname" name="songname" class="pm-input" maxlength="50"
                       value="<%=H(seName)%>" placeholder="Nombre de la cancion" />
            </div>

            <div class="pm-field">
                <label for="genre">Genero</label>
                <input type="text" id="genre" name="genre" class="pm-input" list="lista-generos"
                       value="<%=H(seGenre)%>" placeholder="Rock, Pop, Electronic..." maxlength="50" />
                <datalist id="lista-generos">
                    <option value="Electronic"></option>
                    <option value="Chillout"></option>
                    <option value="Rock"></option>
                    <option value="Pop"></option>
                    <option value="Retro"></option>
                    <option value="Classical"></option>
                    <option value="Country"></option>
                    <option value="Jazz"></option>
                    <option value="Hip-Hop"></option>
                    <option value="Other"></option>
                </datalist>
                <span class="pm-hint">Puedes escribir un genero nuevo o elegir uno existente.</span>
            </div>

            <div class="pm-field">
                <label for="plays">Reproducciones</label>
                <input type="number" id="plays" name="plays" class="pm-input" min="0" value="<%=sePlays%>" />
            </div>

            <div class="pm-field">
                <label for="audio">Archivo de audio</label>
                <input type="file" id="audio" name="audio" class="pm-input" accept="audio/*" />
                <span class="pm-hint">MP3, WAV, OGG, M4A, AAC o FLAC. Se guarda en /SongFiles.</span>
            </div>

            <div class="pm-field pm-span-2">
                <label for="songpath">Ruta del archivo (si no subes uno)</label>
                <input type="text" id="songpath" name="songpath" class="pm-input"
                       value="<%=H(sePath)%>" placeholder="~/SongFiles/bensound-adventure.mp3" />
                <span class="pm-hint">Formato que usa la base de datos: ~/carpeta/archivo.ext</span>
            </div>
        </div>
    </div>

    <div class="pm-panel">
        <h3>Artistas</h3>
        <div class="pm-field">
            <label for="artists">Selecciona los artistas (Ctrl o Cmd para varios)</label>
            <select id="artists" name="artists" class="pm-select" multiple size="10">
                <% Do While Not rsArtistas.EOF
                       Dim aid
                       aid = CStr(rsArtistas("ID_ARTIST")) %>
                <option value="<%=aid%>"<%=IIf(InStr("," & seIds & ",", "," & aid & ",") > 0, " selected=""selected""", "")%>><%=H(NombreArtista(rsArtistas("ARTIST_NAME"), rsArtistas("ARTIST_LASTNAME")))%></option>
                <% rsArtistas.MoveNext
                   Loop %>
            </select>
            <span class="pm-hint">Esta relacion se guarda en TBL_PLAYER.</span>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(seNew, "Crear cancion", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("admin/songs.asp")%>">Cancelar</a>
    </div>
</form>

<% rsArtistas.Close : Set rsArtistas = Nothing %>

<!--#include file="../includes/footer.asp" -->
