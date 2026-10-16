<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar usuario (solo administradores)
' =====================================================================
RequireAdmin

Dim ueId, ueNew, ueUser, ueEmail, ueBirth, ueType, uePass, uePass2, uePhoto, ueErr
Dim rsU, ueUp, ueSelf

ueId    = ToInt(Qry("id"), 0)
ueNew   = (ueId = 0)
ueErr   = ""
ueUser  = ""
ueEmail = ""
ueBirth = ""
ueType  = USER_TYPE_NORMAL
uePhoto = ""

If Not ueNew Then
    Set rsU = Query("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = " & ueId)
    If rsU.EOF Then
        rsU.Close : Set rsU = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    ueUser  = "" & rsU("USERNAME")
    ueEmail = "" & rsU("EMAIL")
    ueBirth = FechaInput(rsU("BIRTHDAY"))
    ueType  = ToInt(rsU("USER_TYPE"), 0)
    uePhoto = "" & rsU("USER_PHOTO")
    rsU.Close : Set rsU = Nothing
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    ueUser  = PostVal("username")
    ueEmail = PostVal("email")
    ueBirth = PostVal("birthday")
    ueType  = ToInt(PostVal("usertype"), USER_TYPE_NORMAL)
    uePass  = PostVal("password")
    uePass2 = PostVal("password2")

    If Len(ueUser) < 3 Then
        ueErr = "El usuario debe tener al menos 3 caracteres."
    ElseIf Not IsValidEmail(ueEmail) Then
        ueErr = "Ingresa un correo electronico valido."
    ElseIf ueNew And Len(uePass) < 6 Then
        ueErr = "La contrasena es obligatoria para un usuario nuevo (minimo 6 caracteres)."
    ElseIf uePass <> "" And Len(uePass) < 6 Then
        ueErr = "La contrasena debe tener al menos 6 caracteres."
    ElseIf uePass <> uePass2 Then
        ueErr = "Las contrasenas no coinciden."
    ElseIf ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = " & SqlText(ueUser) & " AND ID_USER <> " & ueId), 0) > 0 Then
        ueErr = "Ese nombre de usuario ya esta en uso."
    ElseIf ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = " & SqlText(ueEmail) & " AND ID_USER <> " & ueId), 0) > 0 Then
        ueErr = "Ese correo ya esta registrado."
    ElseIf ueId = CurrentUserId() And ueType <> USER_TYPE_ADMIN Then
        ueErr = "No puedes quitarte a ti mismo el rol de administrador."
    End If

    If ueErr = "" Then
        If HasUpload("photo") Then
            ueUp = SaveUpload("photo", DIR_USERS, "|.jpg|.jpeg|.png|.gif|.webp|")
            If ueUp <> "" Then uePhoto = ueUp
        End If
        If uePhoto = "" Then uePhoto = "~/" & DIR_USERS & "/" & IMG_USER_DEFAULT

        If ueNew Then
            ueId = ExecuteInsert("INSERT INTO TBL_USER(USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES(" & _
                   SqlText(ueUser) & ", UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', " & SqlAscii(uePass) & "), 2)), " & _
                   SqlText(ueEmail) & ", " & SqlDate(ueBirth) & ", " & SqlText(uePhoto) & ", " & SqlInt0(ueType) & ")")
            Flash "Usuario """ & ueUser & """ creado correctamente.", "ok"
        Else
            Execute "UPDATE TBL_USER SET USERNAME = " & SqlText(ueUser) & _
                    ", EMAIL = " & SqlText(ueEmail) & _
                    ", BIRTHDAY = " & SqlDate(ueBirth) & _
                    ", USER_PHOTO = " & SqlText(uePhoto) & _
                    ", USER_TYPE = " & SqlInt0(ueType) & _
                    " WHERE ID_USER = " & ueId

            If uePass <> "" Then
                Execute "UPDATE TBL_USER SET PASSWORD = UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', " & SqlAscii(uePass) & "), 2)) WHERE ID_USER = " & ueId
            End If

            ' Si el administrador se edito a si mismo, refresca la sesion.
            If ueId = CurrentUserId() Then RefreshSession

            Flash "Usuario """ & ueUser & """ actualizado.", "ok"
        End If

        Response.Redirect AppUrl("admin/users.asp")
        Response.End
    End If
End If

ueSelf = (ueId = CurrentUserId())

gPage = "admin-users"
gTitle = IIf(ueNew, "Nuevo usuario", "Editar usuario")
%>
<!--#include file="../includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/users.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a usuarios
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub"><%=IIf(ueNew, "Crea una cuenta y define su rol en el sistema.", "Modifica los datos, el rol o la contrasena del usuario.")%></p>

<% If ueErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(ueErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("admin/user_edit.asp")%>?id=<%=ueId%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Datos de la cuenta</h3>

        <div class="pm-form-grid">
            <div class="pm-field">
                <label for="username">Usuario</label>
                <input type="text" id="username" name="username" class="pm-input" value="<%=H(ueUser)%>" placeholder="nombre_usuario" />
            </div>

            <div class="pm-field">
                <label for="email">Correo electronico</label>
                <input type="email" id="email" name="email" class="pm-input" value="<%=H(ueEmail)%>" placeholder="correo@epn.edu.ec" />
            </div>

            <div class="pm-field">
                <label for="birthday">Fecha de nacimiento</label>
                <input type="date" id="birthday" name="birthday" class="pm-input" value="<%=H(ueBirth)%>" />
            </div>

            <div class="pm-field">
                <label for="usertype">Tipo de usuario</label>
                <select id="usertype" name="usertype" class="pm-select">
                    <option value="0"<%=SelIf(USER_TYPE_NORMAL, ueType)%>>Usuario</option>
                    <option value="1"<%=SelIf(USER_TYPE_ADMIN, ueType)%>>Administrador</option>
                </select>
                <span class="pm-hint">El administrador tiene acceso total al CRUD del sistema.</span>
            </div>

            <div class="pm-field">
                <label for="password">Contrasena</label>
                <input type="password" id="password" name="password" class="pm-input"
                       placeholder="<%=IIf(ueNew, "Minimo 6 caracteres", "Dejar vacio para no cambiarla")%>"
                       autocomplete="new-password" />
            </div>

            <div class="pm-field">
                <label for="password2">Confirmar contrasena</label>
                <input type="password" id="password2" name="password2" class="pm-input" autocomplete="new-password" />
            </div>

            <div class="pm-field">
                <label for="photo">Foto de perfil</label>
                <input type="file" id="photo" name="photo" class="pm-input" accept="image/*" />
                <span class="pm-hint">JPG, PNG, GIF o WEBP.</span>
            </div>

            <% If uePhoto <> "" Then %>
            <div class="pm-field">
                <label>Foto actual</label>
                <img src="<%=H(MediaUrl(uePhoto, "user", ueUser))%>" alt="Foto"
                     style="width:90px;height:90px;border-radius:50%;object-fit:cover;border:1px solid var(--stroke)" />
            </div>
            <% End If %>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(ueNew, "Crear usuario", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("admin/users.asp")%>">Cancelar</a>
    </div>
</form>

<!--#include file="../includes/footer.asp" -->
