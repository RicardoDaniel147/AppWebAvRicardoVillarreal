<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar album con su tracklist (solo administradores)
' =====================================================================
RequireAdmin

Dim leId, leNew, leTitle, leYear, leGenre, leImage, leErr, leSongs, leIds
Dim rsL, rsSongs, leUp, leArr, lei

leId    = ToInt(Qry("id"), 0)
leNew   = (leId = 0)
leErr   = ""
leTitle = ""
leYear  = Year(Date())
leGenre = ""
leImage = ""
leIds   = ""

If Not leNew Then
    Set rsL = Query("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH, GENRE FROM TBL_ALBUM WHERE ID_ALBUM = " & leId)
    If rsL.EOF Then
        rsL.Close : Set rsL = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    leTitle = "" & rsL("TITLE_ALBUM")
    If Not IsNull(rsL("RELEASE_YEAR")) Then leYear = ToInt(rsL("RELEASE_YEAR"), 0)
    leGenre = "" & rsL("GENRE")
    leImage = "" & rsL("ALBUM_IMAGE_PATH")
    rsL.Close : Set rsL = Nothing

    Set rsL = Query("SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = " & leId)
    Do While Not rsL.EOF
        leIds = leIds & IIf(leIds = "", "", ",") & rsL("ID_SONG")
        rsL.MoveNext
    Loop
    rsL.Close : Set rsL = Nothing
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    leTitle = PostVal("titlealbum")
    leYear  = ToInt(PostVal("releaseyear"), 0)
    leGenre = PostVal("genre")
    leIds   = PostVal("songs")

    If Len(Trim(leTitle)) = 0 Then
        leErr = "El titulo del album es obligatorio."
    ElseIf leYear > 0 And (leYear < 1900 Or leYear > 2100) Then
        leErr = "El ano de lanzamiento no es valido."
    End If

    If leErr = "" Then
        If HasUpload("image") Then
            leUp = SaveUpload("image", DIR_ALBUMS, "|.jpg|.jpeg|.png|.gif|.webp|")
            If leUp <> "" Then leImage = leUp
        End If

        If leNew Then
            leId = ExecuteInsert("INSERT INTO TBL_ALBUM(TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH, GENRE) VALUES(" & _
                   SqlText(Trim(leTitle)) & ", " & SqlInt(IIf(leYear > 0, leYear, "")) & ", " & _
                   SqlTextOrNull(leImage) & ", " & SqlTextOrNull(leGenre) & ")")
            Flash "Album """ & Trim(leTitle) & """ creado correctamente.", "ok"
        Else
            Execute "UPDATE TBL_ALBUM SET TITLE_ALBUM = " & SqlText(Trim(leTitle)) & _
                    ", RELEASE_YEAR = " & SqlInt(IIf(leYear > 0, leYear, "")) & _
                    ", ALBUM_IMAGE_PATH = " & SqlTextOrNull(leImage) & _
                    ", GENRE = " & SqlTextOrNull(leGenre) & _
                    " WHERE ID_ALBUM = " & leId
            Execute "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & leId
            Flash "Album """ & Trim(leTitle) & """ actualizado.", "ok"
        End If

        If leIds <> "" Then
            leArr = Split(leIds, ",")
            For lei = 0 To UBound(leArr)
                If ToInt(leArr(lei), 0) > 0 Then
                    Execute "INSERT INTO TBL_TRACKLIST(ID_ALBUM, ID_SONG) VALUES(" & leId & ", " & ToInt(leArr(lei), 0) & ")"
                End If
            Next
        End If

        Response.Redirect AppUrl("admin/albums.asp")
        Response.End
    End If
End If

Set rsSongs = Query(SqlCanciones("", "s.SONG_NAME", 0))

gPage = "admin-albums"
gTitle = IIf(leNew, "Nuevo album", "Editar album")
%>
<!--#include file="../includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/albums.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a albumes
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub"><%=IIf(leNew, "Crea un album y selecciona sus canciones.", "Actualiza los datos y el tracklist del album.")%></p>

<% If leErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(leErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("admin/album_edit.asp")%>?id=<%=leId%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Datos del album</h3>

        <div class="pm-form-grid">
            <div class="pm-field pm-span-2">
                <label for="titlealbum">Titulo del album</label>
                <input type="text" id="titlealbum" name="titlealbum" class="pm-input" maxlength="50"
                       value="<%=H(leTitle)%>" placeholder="Titulo del album" />
            </div>

            <div class="pm-field">
                <label for="genre">Genero</label>
                <input type="text" id="genre" name="genre" class="pm-input" list="lista-generos-album"
                       value="<%=H(leGenre)%>" placeholder="Rock, Pop, Classical..." maxlength="50" />
                <datalist id="lista-generos-album">
                    <option value="Rock"></option><option value="Chillout"></option>
                    <option value="Retro"></option><option value="Classical"></option>
                    <option value="Country"></option><option value="Electronic"></option>
                    <option value="Pop"></option><option value="Jazz"></option>
                </datalist>
            </div>

            <div class="pm-field">
                <label for="releaseyear">Ano de lanzamiento</label>
                <input type="number" id="releaseyear" name="releaseyear" class="pm-input" min="1900" max="2100" value="<%=leYear%>" />
            </div>

            <div class="pm-field">
                <label for="image">Portada del album</label>
                <input type="file" id="image" name="image" class="pm-input" accept="image/*" />
                <span class="pm-hint">Se guarda en /AlbumImages.</span>
            </div>

            <% If leImage <> "" Then %>
            <div class="pm-field">
                <label>Portada actual</label>
                <img src="<%=H(MediaUrl(leImage, "album", leTitle))%>" alt="Portada"
                     style="width:110px;height:110px;border-radius:12px;object-fit:cover;border:1px solid var(--stroke)" />
            </div>
            <% End If %>
        </div>
    </div>

    <div class="pm-panel">
        <h3>Tracklist</h3>
        <div class="pm-field">
            <label for="songs">Selecciona las canciones del album (Ctrl o Cmd para varias)</label>
            <select id="songs" name="songs" class="pm-select" multiple size="14">
                <% Do While Not rsSongs.EOF
                       Dim lsid
                       lsid = CStr(rsSongs("ID_SONG")) %>
                <option value="<%=lsid%>"<%=IIf(InStr("," & leIds & ",", "," & lsid & ",") > 0, " selected=""selected""", "")%>><%=H(rsSongs("SONG_NAME"))%> &mdash; <%=H("" & rsSongs("GENRE"))%></option>
                <% rsSongs.MoveNext
                   Loop %>
            </select>
            <span class="pm-hint">Esta relacion se guarda en TBL_TRACKLIST.</span>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(leNew, "Crear album", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("admin/albums.asp")%>">Cancelar</a>
    </div>
</form>

<% rsSongs.Close : Set rsSongs = Nothing %>

<!--#include file="../includes/footer.asp" -->
