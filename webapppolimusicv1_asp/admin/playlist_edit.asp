<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar playlist y reasignar propietario (admin)
' =====================================================================
RequireAdmin

Dim lpId, lpNew, lpTitle, lpType, lpDate, lpPhoto, lpOwner, lpErr, lpIds
Dim rsP, rsSongs, rsUsers, lpUp, lpArr, lpi

lpId    = ToInt(Qry("id"), 0)
lpNew   = (lpId = 0)
lpErr   = ""
lpTitle = ""
lpType  = PLAYLIST_PUBLIC
lpDate  = Date()
lpPhoto = ""
lpOwner = CurrentUserId()
lpIds   = ""

If Not lpNew Then
    Set rsP = Query("SELECT ID_PLAYLIST, ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & lpId)
    If rsP.EOF Then
        rsP.Close : Set rsP = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    lpOwner = ToInt(rsP("ID_USER"), 0)
    lpTitle = "" & rsP("TITLE")
    lpType  = ToInt(rsP("TYPE"), PLAYLIST_PUBLIC)
    lpDate  = rsP("CREATION_DATE")
    lpPhoto = "" & rsP("PHOTO")
    rsP.Close : Set rsP = Nothing

    Set rsP = Query("SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & lpId)
    Do While Not rsP.EOF
        lpIds = lpIds & IIf(lpIds = "", "", ",") & rsP("ID_SONG")
        rsP.MoveNext
    Loop
    rsP.Close : Set rsP = Nothing
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    lpTitle = PostVal("title")
    lpType  = ToInt(PostVal("type"), PLAYLIST_PUBLIC)
    lpDate  = PostVal("creationdate")
    lpOwner = ToInt(PostVal("iduser"), 0)
    lpIds   = PostVal("songs")
    If lpDate = "" Then lpDate = Date()

    If Len(Trim(lpTitle)) = 0 Then
        lpErr = "El titulo es obligatorio."
    ElseIf lpOwner <= 0 Then
        lpErr = "Selecciona un propietario valido."
    End If

    If lpErr = "" Then
        If HasUpload("photo") Then
            lpUp = SaveUpload("photo", DIR_PLAYLISTS, "|.jpg|.jpeg|.png|.gif|.webp|")
            If lpUp <> "" Then lpPhoto = lpUp
        End If
        If lpPhoto = "" Then lpPhoto = "~/" & DIR_PLAYLISTS & "/" & IMG_PLAYLIST_DEFAULT

        If lpNew Then
            lpId = ExecuteInsert("INSERT INTO TBL_PLAYLIST(ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES(" & _
                   SqlInt0(lpOwner) & ", " & SqlText(Trim(lpTitle)) & ", " & SqlDate(lpDate) & ", " & _
                   SqlInt0(lpType) & ", " & SqlText(lpPhoto) & ")")
            Flash "Playlist """ & Trim(lpTitle) & """ creada correctamente.", "ok"
        Else
            Execute "UPDATE TBL_PLAYLIST SET ID_USER = " & SqlInt0(lpOwner) & _
                    ", TITLE = " & SqlText(Trim(lpTitle)) & _
                    ", CREATION_DATE = " & SqlDate(lpDate) & _
                    ", TYPE = " & SqlInt0(lpType) & _
                    ", PHOTO = " & SqlText(lpPhoto) & _
                    " WHERE ID_PLAYLIST = " & lpId
            Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & lpId
            Flash "Playlist """ & Trim(lpTitle) & """ actualizada.", "ok"
        End If

        If lpIds <> "" Then
            lpArr = Split(lpIds, ",")
            For lpi = 0 To UBound(lpArr)
                If ToInt(lpArr(lpi), 0) > 0 Then
                    Execute "INSERT INTO TBL_PLAYLIST_SONG(ID_PLAYLIST, ID_SONG) VALUES(" & lpId & ", " & ToInt(lpArr(lpi), 0) & ")"
                End If
            Next
        End If

        Response.Redirect AppUrl("admin/playlists.asp")
        Response.End
    End If
End If

Set rsUsers = Query("SELECT ID_USER, USERNAME, USER_TYPE FROM TBL_USER ORDER BY USERNAME")
Set rsSongs = Query(SqlCanciones("", "s.SONG_NAME", 0))

gPage = "admin-playlists"
gTitle = IIf(lpNew, "Nueva playlist", "Editar playlist")
%>
<!--#include file="../includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/playlists.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a playlists
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub"><%=IIf(lpNew, "Crea una playlist y asignala a un usuario.", "Modifica la playlist, su propietario y sus canciones.")%></p>

<% If lpErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(lpErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("admin/playlist_edit.asp")%>?id=<%=lpId%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Datos de la playlist</h3>

        <div class="pm-form-grid">
            <div class="pm-field pm-span-2">
                <label for="title">Titulo</label>
                <input type="text" id="title" name="title" class="pm-input" maxlength="50"
                       value="<%=H(lpTitle)%>" placeholder="Titulo de la playlist" />
            </div>

            <div class="pm-field">
                <label for="iduser">Propietario</label>
                <select id="iduser" name="iduser" class="pm-select">
                    <% Do While Not rsUsers.EOF %>
                    <option value="<%=rsUsers("ID_USER")%>"<%=IIf(lpOwner = ToInt(rsUsers("ID_USER"), 0), " selected=""selected""", "")%>><%=H(rsUsers("USERNAME"))%> (<%=TipoUsuarioNombre(rsUsers("USER_TYPE"))%>)</option>
                    <% rsUsers.MoveNext
                       Loop %>
                </select>
            </div>

            <div class="pm-field">
                <label for="type">Tipo</label>
                <select id="type" name="type" class="pm-select">
                    <option value="0"<%=SelIf(PLAYLIST_PUBLIC, lpType)%>>Publica</option>
                    <option value="1"<%=SelIf(PLAYLIST_PRIVATE, lpType)%>>Privada</option>
                </select>
            </div>

            <div class="pm-field">
                <label for="creationdate">Fecha de creacion</label>
                <input type="date" id="creationdate" name="creationdate" class="pm-input" value="<%=FechaInput(lpDate)%>" />
            </div>

            <div class="pm-field">
                <label for="photo">Portada</label>
                <input type="file" id="photo" name="photo" class="pm-input" accept="image/*" />
                <span class="pm-hint">Se guarda en /PlaylistImages.</span>
            </div>

            <% If lpPhoto <> "" Then %>
            <div class="pm-field">
                <label>Portada actual</label>
                <img src="<%=H(MediaUrl(lpPhoto, "playlist", lpTitle))%>" alt="Portada"
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
                       Dim psid
                       psid = CStr(rsSongs("ID_SONG")) %>
                <option value="<%=psid%>"<%=IIf(InStr("," & lpIds & ",", "," & psid & ",") > 0, " selected=""selected""", "")%>><%=H(rsSongs("SONG_NAME"))%> &mdash; <%=H("" & rsSongs("GENRE"))%></option>
                <% rsSongs.MoveNext
                   Loop %>
            </select>
            <span class="pm-hint">Esta relacion se guarda en TBL_PLAYLIST_SONG.</span>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(lpNew, "Crear playlist", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("admin/playlists.asp")%>">Cancelar</a>
    </div>
</form>

<% rsUsers.Close : Set rsUsers = Nothing %>
<% rsSongs.Close : Set rsSongs = Nothing %>

<!--#include file="../includes/footer.asp" -->
