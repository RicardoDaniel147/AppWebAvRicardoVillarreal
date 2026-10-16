<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar artista (solo administradores)
' =====================================================================
RequireAdmin

Dim aeId, aeNew, aeName, aeLast, aeImage, aeErr, rsA, aeUp

aeId    = ToInt(Qry("id"), 0)
aeNew   = (aeId = 0)
aeErr   = ""
aeName  = ""
aeLast  = ""
aeImage = ""

If Not aeNew Then
    Set rsA = Query("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = " & aeId)
    If rsA.EOF Then
        rsA.Close : Set rsA = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    aeName  = "" & rsA("ARTIST_NAME")
    aeLast  = "" & rsA("ARTIST_LASTNAME")
    aeImage = "" & rsA("ARTIST_IMAGE")
    rsA.Close : Set rsA = Nothing
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    aeName = PostVal("artistname")
    aeLast = PostVal("artistlastname")

    If Len(Trim(aeName)) = 0 Then
        aeErr = "El nombre del artista es obligatorio."
    End If

    If aeErr = "" Then
        If HasUpload("image") Then
            aeUp = SaveUpload("image", DIR_ARTISTS, "|.jpg|.jpeg|.png|.gif|.webp|")
            If aeUp <> "" Then aeImage = aeUp
        End If

        If aeNew Then
            aeId = ExecuteInsert("INSERT INTO TBL_ARTIST(ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES(" & _
                   SqlText(Trim(aeName)) & ", " & SqlTextOrNull(Trim(aeLast)) & ", " & SqlTextOrNull(aeImage) & ")")
            Flash "Artista """ & Trim(aeName) & """ creado correctamente.", "ok"
        Else
            Execute "UPDATE TBL_ARTIST SET ARTIST_NAME = " & SqlText(Trim(aeName)) & _
                    ", ARTIST_LASTNAME = " & SqlTextOrNull(Trim(aeLast)) & _
                    ", ARTIST_IMAGE = " & SqlTextOrNull(aeImage) & _
                    " WHERE ID_ARTIST = " & aeId
            Flash "Artista """ & Trim(aeName) & """ actualizado.", "ok"
        End If

        Response.Redirect AppUrl("admin/artists.asp")
        Response.End
    End If
End If

gPage = "admin-artists"
gTitle = IIf(aeNew, "Nuevo artista", "Editar artista")
%>
<!--#include file="../includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/artists.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a artistas
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub"><%=IIf(aeNew, "Registra un nuevo artista y su imagen.", "Actualiza los datos del artista.")%></p>

<% If aeErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(aeErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("admin/artist_edit.asp")%>?id=<%=aeId%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Datos del artista</h3>

        <div class="pm-form-grid">
            <div class="pm-field">
                <label for="artistname">Nombre</label>
                <input type="text" id="artistname" name="artistname" class="pm-input" maxlength="100"
                       value="<%=H(aeName)%>" placeholder="Nombre" />
            </div>

            <div class="pm-field">
                <label for="artistlastname">Apellido</label>
                <input type="text" id="artistlastname" name="artistlastname" class="pm-input" maxlength="100"
                       value="<%=H(aeLast)%>" placeholder="Apellido" />
            </div>

            <div class="pm-field">
                <label for="image">Imagen del artista</label>
                <input type="file" id="image" name="image" class="pm-input" accept="image/*" />
                <span class="pm-hint">Se guarda en /ArtistImages.</span>
            </div>

            <% If aeImage <> "" Then %>
            <div class="pm-field">
                <label>Imagen actual</label>
                <img src="<%=H(MediaUrl(aeImage, "artist", NombreArtista(aeName, aeLast)))%>" alt="Imagen"
                     style="width:96px;height:96px;border-radius:50%;object-fit:cover;border:1px solid var(--stroke)" />
            </div>
            <% End If %>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(aeNew, "Crear artista", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("admin/artists.asp")%>">Cancelar</a>
    </div>
</form>

<!--#include file="../includes/footer.asp" -->
