<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar una playlist (dueno o administrador)
' =====================================================================
RequireLogin

Dim peId, peNew, peTitle, peType, peDate, pePhoto, peErr, peOwner
Dim rsPl, rsSongs, peSelSongs, peUp, pi, peSongIds

peId    = ToInt(Qry("id"), 0)
peNew   = (peId = 0)
peErr   = ""
peTitle = ""
peType  = PLAYLIST_PUBLIC
peDate  = Date()
pePhoto = ""
peOwner = CurrentUserId()
peSongIds = ""

If Not peNew Then
    Set rsPl = Query(SqlPlaylists("pl.ID_PLAYLIST = " & peId, "", 0))
    If rsPl.EOF Then
        rsPl.Close : Set rsPl = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    If Not PuedeEditarPlaylist(peId, ToInt(rsPl("ID_USER"), 0), ToInt(rsPl("TYPE"), 0)) Then
        rsPl.Close : Set rsPl = Nothing
        Response.Redirect AppUrl("error.asp?code=403")
        Response.End
    End If
    peTitle = "" & rsPl("TITLE")
    peType  = ToInt(rsPl("TYPE"), PLAYLIST_PUBLIC)
    peDate  = rsPl("CREATION_DATE")
    pePhoto = "" & rsPl("PHOTO")
    peOwner = ToInt(rsPl("ID_USER"), CurrentUserId())
    rsPl.Close : Set rsPl = Nothing
End If

' ---------------------------------------------------------------- guardar
If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    peTitle = PostVal("title")
    peType  = ToInt(PostVal("type"), PLAYLIST_PUBLIC)
    peDate  = PostVal("creationdate")
    If peDate = "" Then peDate = Date()
    peSongIds = PostVal("songs")

    If Len(Trim(peTitle)) = 0 Then
        peErr = "El titulo es obligatorio."
    ElseIf Len(Trim(peTitle)) > 50 Then
        peErr = "El titulo no puede superar los 50 caracteres."
    ElseIf Not IsAdmin() Then
        peOwner = CurrentUserId()   ' un usuario normal no puede reasignar la playlist
    End If

    If peErr = "" Then
        If HasUpload("photo") Then
            peUp = SaveUpload("photo", DIR_PLAYLISTS, "|.jpg|.jpeg|.png|.gif|.webp|")
            If peUp <> "" Then pePhoto = peUp
        End If
        If pePhoto = "" Then pePhoto = "~/" & DIR_PLAYLISTS & "/" & IMG_PLAYLIST_DEFAULT

        If peNew Then
            peId = ExecuteInsert("INSERT INTO TBL_PLAYLIST(ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES(" & _
                   SqlInt0(peOwner) & ", " & SqlText(Trim(peTitle)) & ", " & SqlDate(peDate) & ", " & _
                   SqlInt0(peType) & ", " & SqlText(pePhoto) & ")")
        Else
            Execute "UPDATE TBL_PLAYLIST SET ID_USER = " & SqlInt0(peOwner) & _
                    ", TITLE = " & SqlText(Trim(peTitle)) & _
                    ", CREATION_DATE = " & SqlDate(peDate) & _
                    ", TYPE = " & SqlInt0(peType) & _
                    ", PHOTO = " & SqlText(pePhoto) & _
                    " WHERE ID_PLAYLIST = " & peId
            Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & peId
        End If

        ' Canciones seleccionadas
        If peSongIds <> "" Then
            Dim arrIds
            arrIds = Split(peSongIds, ",")
            For pi = 0 To UBound(arrIds)
                If ToInt(arrIds(pi), 0) > 0 Then
                    Execute "INSERT INTO TBL_PLAYLIST_SONG(ID_PLAYLIST, ID_SONG) VALUES(" & peId & ", " & ToInt(arrIds(pi), 0) & ")"
                End If
            Next
        End If

        Flash "Playlist """ & Trim(peTitle) & """ guardada correctamente.", "ok"
        Response.Redirect AppUrl("playlist.asp") & "?id=" & peId
        Response.End
    End If
End If

' Canciones ya seleccionadas (para marcar el select multiple)
If Not peNew Then
    Set rsSongs = Query("SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & peId)
    Do While Not rsSongs.EOF
        peSongIds = peSongIds & IIf(peSongIds = "", "", ",") & rsSongs("ID_SONG")
        rsSongs.MoveNext
    Loop
    rsSongs.Close : Set rsSongs = Nothing
End If

Set rsSongs = Query(SqlCanciones("", "s.SONG_NAME", 0))

gPage = "playlists"
gTitle = IIf(peNew, "Nueva playlist", "Editar playlist")
%>
<!--#include file="includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("playlists.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub"><%=IIf(peNew, "Crea una coleccion con tus canciones favoritas.", "Actualiza los datos y las canciones de la playlist.")%></p>

<% If peErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(peErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("playlist_edit.asp")%>?id=<%=peId%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Datos generales</h3>

        <div class="pm-form-grid">
            <div class="pm-field pm-span-2">
                <label for="title">Titulo</label>
                <input type="text" id="title" name="title" class="pm-input" maxlength="50"
                       value="<%=H(peTitle)%>" placeholder="Mi playlist favorita" />
            </div>

            <div class="pm-field">
                <label for="type">Tipo</label>
                <select id="type" name="type" class="pm-select">
                    <option value="0"<%=SelIf(PLAYLIST_PUBLIC, peType)%>>Publica</option>
                    <option value="1"<%=SelIf(PLAYLIST_PRIVATE, peType)%>>Privada</option>
                </select>
                <span class="pm-hint">Las playlists privadas solo las ves tu.</span>
            </div>

            <div class="pm-field">
                <label for="creationdate">Fecha de creacion</label>
                <input type="date" id="creationdate" name="creationdate" class="pm-input" value="<%=FechaInput(peDate)%>" />
            </div>

            <div class="pm-field">
                <label for="photo">Portada</label>
                <input type="file" id="photo" name="photo" class="pm-input" accept="image/*" />
                <span class="pm-hint">JPG, PNG, GIF o WEBP. Si no subes nada se conserva la actual.</span>
            </div>

            <% If pePhoto <> "" Then %>
            <div class="pm-field">
                <label>Portada actual</label>
                <img src="<%=H(MediaUrl(pePhoto, "playlist", peTitle))%>" alt="Portada"
                     style="width:110px;height:110px;border-radius:12px;object-fit:cover;border:1px solid var(--stroke)" />
            </div>
            <% End If %>
        </div>
    </div>

    <div class="pm-panel">
        <h3>Canciones</h3>
        <div class="pm-field">
            <label for="songs">Selecciona las canciones (Ctrl o Cmd para varias)</label>
            <select id="songs" name="songs" class="pm-select" multiple size="14">
                <% Do While Not rsSongs.EOF
                       Dim sid
                       sid = CStr(rsSongs("ID_SONG")) %>
                <option value="<%=sid%>"<%=IIf(InStr("," & peSongIds & ",", "," & sid & ",") > 0, " selected=""selected""", "")%>><%=H(rsSongs("SONG_NAME"))%> &mdash; <%=H("" & rsSongs("GENRE"))%></option>
                <% rsSongs.MoveNext
                   Loop %>
            </select>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(peNew, "Crear playlist", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("playlists.asp")%>">Cancelar</a>
    </div>
</form>

<% rsSongs.Close : Set rsSongs = Nothing %>

<!--#include file="includes/footer.asp" -->
