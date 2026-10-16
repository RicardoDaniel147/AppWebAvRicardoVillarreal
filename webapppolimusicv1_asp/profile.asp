<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Perfil del usuario: ver y editar
' =====================================================================
RequireLogin

Dim pfEdit, pfUser, pfEmail, pfBirth, pfErr, pfOk, pfPhoto, pfUp, pfNewPass, pfPass2
Dim rsUser, rsMine

pfEdit = (LCase(Qry("action")) = "edit")
pfErr  = ""
pfPhoto = CurrentUserPhoto()
pfUser  = CurrentUsername()
pfEmail = CurrentUserEmail()
pfBirth = ""

Set rsUser = Query("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = " & CurrentUserId())
If rsUser.EOF Then
    rsUser.Close : Set rsUser = Nothing
    Response.Redirect AppUrl("error.asp?code=404")
    Response.End
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    pfEdit  = True
    pfUser  = PostVal("username")
    pfEmail = PostVal("email")
    pfBirth = PostVal("birthday")
    pfNewPass = PostVal("newpassword")
    pfPass2   = PostVal("password2")

    If Len(pfUser) < 3 Then
        pfErr = "El usuario debe tener al menos 3 caracteres."
    ElseIf Not IsValidEmail(pfEmail) Then
        pfErr = "Ingresa un correo electronico valido."
    ElseIf ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = " & SqlText(pfUser) & " AND ID_USER <> " & CurrentUserId()), 0) > 0 Then
        pfErr = "Ese nombre de usuario ya esta en uso."
    ElseIf ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = " & SqlText(pfEmail) & " AND ID_USER <> " & CurrentUserId()), 0) > 0 Then
        pfErr = "Ese correo ya esta registrado."
    ElseIf pfNewPass <> "" And Len(pfNewPass) < 6 Then
        pfErr = "La nueva contrasena debe tener al menos 6 caracteres."
    ElseIf pfNewPass <> pfPass2 Then
        pfErr = "Las contrasenas no coinciden."
    End If

    If pfErr = "" Then
        If HasUpload("photo") Then
            pfUp = SaveUpload("photo", DIR_USERS, "|.jpg|.jpeg|.png|.gif|.webp|")
            If pfUp <> "" Then pfPhoto = pfUp
        End If

        Execute "UPDATE TBL_USER SET USERNAME = " & SqlText(pfUser) & _
                ", EMAIL = " & SqlText(pfEmail) & _
                ", BIRTHDAY = " & SqlDate(pfBirth) & _
                ", USER_PHOTO = " & SqlText(pfPhoto) & _
                " WHERE ID_USER = " & CurrentUserId()

        If pfNewPass <> "" Then
            Execute "UPDATE TBL_USER SET PASSWORD = UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', " & SqlAscii(pfNewPass) & "), 2)) WHERE ID_USER = " & CurrentUserId()
        End If

        RefreshSession
        Flash "Perfil actualizado correctamente.", "ok"
        Response.Redirect AppUrl("profile.asp")
        Response.End
    End If
End If

pfBirth = FechaInput(rsUser("BIRTHDAY"))

Set rsMine = Query(SqlPlaylists("pl.ID_USER = " & CurrentUserId(), "pl.CREATION_DATE DESC", 0))

gPage = "profile"
gTitle = "Mi perfil"
%>
<!--#include file="includes/header.asp" -->

<% If Not pfEdit Then %>

<section class="pm-detail pm-fade-in">
    <div class="pm-detail-art round">
        <img src="<%=H(MediaUrl(pfPhoto, "user", CurrentUsername()))%>" alt="<%=H(CurrentUsername())%>" />
    </div>

    <div>
        <span class="pm-eyebrow">Perfil</span>
        <h1><%=H(rsUser("USERNAME"))%></h1>
        <div class="pm-detail-meta">
            <span class="<%=BadgeUsuario(ToInt(rsUser("USER_TYPE"), 0) = USER_TYPE_ADMIN)%>"><%=TipoUsuarioNombre(rsUser("USER_TYPE"))%></span>
            <span><%=H(rsUser("EMAIL"))%></span>
            <% If Not IsNull(rsUser("BIRTHDAY")) Then %>
            <span class="dot"></span>
            <span>Nacimiento: <%=FechaCorta(rsUser("BIRTHDAY"))%></span>
            <% End If %>
        </div>

        <div class="pm-btn-row" style="margin-top:22px">
            <a class="pm-btn pm-btn-primary" href="<%=AppUrl("profile.asp")%>?action=edit">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M12 20h9" /><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z" />
                </svg>
                Editar perfil
            </a>
            <a class="pm-btn" href="<%=AppUrl("playlists.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <path d="M4 6h16M4 12h16M4 18h10" />
                </svg>
                Mis playlists
            </a>
        </div>
    </div>
</section>

<div class="pm-section">
    <h2>Mis playlists</h2>
    <a href="<%=AppUrl("playlist_edit.asp")%>">Crear nueva</a>
</div>

<% If rsMine.EOF Then %>
<div class="pm-empty"><strong>Sin playlists</strong>Aun no has creado ninguna playlist.</div>
<% Else %>
<div class="pm-grid">
    <% Do While Not rsMine.EOF %>
    <a class="pm-card" href="<%=AppUrl("playlist.asp")%>?id=<%=rsMine("ID_PLAYLIST")%>">
        <div class="pm-card-art">
            <img src="<%=H(MediaUrl("" & rsMine("PHOTO"), "playlist", "" & rsMine("TITLE")))%>" alt="" loading="lazy" />
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsMine("TITLE"))%></div>
            <div class="pm-card-meta"><%=rsMine("CANCIONES")%> canciones &middot; <%=TipoPlaylistNombre(rsMine("TYPE"))%></div>
        </div>
    </a>
    <% rsMine.MoveNext
       Loop %>
</div>
<% End If %>

<% Else %>

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("profile.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver al perfil
</a>

<h1 class="pm-h1">Editar perfil</h1>
<p class="pm-sub">Actualiza tus datos personales, tu foto o tu contrasena.</p>

<% If pfErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(pfErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("profile.asp")%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-two-col">
        <div class="pm-panel">
            <h3>Datos personales</h3>

            <div class="pm-form-grid" style="grid-template-columns:1fr">
                <div class="pm-field">
                    <label for="username">Usuario</label>
                    <input type="text" id="username" name="username" class="pm-input" value="<%=H(pfUser)%>" />
                </div>

                <div class="pm-field">
                    <label for="email">Correo electronico</label>
                    <input type="email" id="email" name="email" class="pm-input" value="<%=H(pfEmail)%>" />
                </div>

                <div class="pm-field">
                    <label for="birthday">Fecha de nacimiento</label>
                    <input type="date" id="birthday" name="birthday" class="pm-input" value="<%=H(pfBirth)%>" />
                </div>

                <div class="pm-field">
                    <label for="photo">Foto de perfil</label>
                    <input type="file" id="photo" name="photo" class="pm-input" accept="image/*" />
                    <span class="pm-hint">Si no subes una nueva, se mantiene la actual.</span>
                </div>
            </div>
        </div>

        <div class="pm-panel">
            <h3>Cambiar contrasena</h3>
            <p class="pm-hint" style="margin-bottom:16px">Deja estos campos vacios si no quieres cambiarla.</p>

            <div class="pm-form-grid" style="grid-template-columns:1fr">
                <div class="pm-field">
                    <label for="newpassword">Nueva contrasena</label>
                    <input type="password" id="newpassword" name="newpassword" class="pm-input"
                           placeholder="Minimo 6 caracteres" autocomplete="new-password" />
                </div>

                <div class="pm-field">
                    <label for="password2">Confirmar contrasena</label>
                    <input type="password" id="password2" name="password2" class="pm-input" autocomplete="new-password" />
                </div>

                <div class="pm-field">
                    <label>Foto actual</label>
                    <img src="<%=H(MediaUrl(pfPhoto, "user", pfUser))%>" alt="Foto"
                         style="width:96px;height:96px;border-radius:50%;object-fit:cover;border:1px solid var(--stroke)" />
                </div>
            </div>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            Guardar cambios
        </button>
        <a class="pm-btn" href="<%=AppUrl("profile.asp")%>">Cancelar</a>
    </div>
</form>

<% End If %>

<% rsMine.Close : Set rsMine = Nothing %>
<% rsUser.Close : Set rsUser = Nothing %>

<!--#include file="includes/footer.asp" -->
