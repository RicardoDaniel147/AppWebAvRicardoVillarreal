<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Registro de usuarios
'  Todo registro publico crea SIEMPRE un usuario normal (USER_TYPE = 0).
' =====================================================================
Dim rUser, rEmail, rBirth, rPass, rPass2, rError, rsNew, rPhoto, rNewId

rError = ""
rUser = ""
rEmail = ""
rBirth = ""

If IsLoggedIn() Then Response.Redirect AppUrl("home.asp")

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    rUser  = PostVal("username")
    rEmail = PostVal("email")
    rBirth = PostVal("birthday")
    rPass  = PostVal("password")
    rPass2 = PostVal("password2")

    If Len(rUser) < 3 Then
        rError = "El usuario debe tener al menos 3 caracteres."
    ElseIf Not IsValidEmail(rEmail) Then
        rError = "Ingresa un correo electronico valido."
    ElseIf Len(rPass) < 6 Then
        rError = "La contrasena debe tener al menos 6 caracteres."
    ElseIf rPass <> rPass2 Then
        rError = "Las contrasenas no coinciden."
    ElseIf ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = " & SqlText(rUser)), 0) > 0 Then
        rError = "Ese nombre de usuario ya esta en uso."
    ElseIf ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = " & SqlText(rEmail)), 0) > 0 Then
        rError = "Ese correo ya esta registrado."
    Else
        rPhoto = ""
        If HasUpload("photo") Then
            rPhoto = SaveUpload("photo", DIR_USERS, "|.jpg|.jpeg|.png|.gif|.webp|")
        End If
        If rPhoto = "" Then rPhoto = "~/" & DIR_USERS & "/" & IMG_USER_DEFAULT

        rNewId = ExecuteInsert("INSERT INTO TBL_USER(USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES(" & _
                  SqlText(rUser) & ", " & _
                  "UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', " & SqlAscii(rPass) & "), 2)), " & _
                  SqlText(rEmail) & ", " & SqlDate(rBirth) & ", " & _
                  SqlText(rPhoto) & ", " & USER_TYPE_NORMAL & ")")

        If rNewId > 0 Then
            Set rsNew = Query("SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE ID_USER = " & rNewId)
            If Not rsNew.EOF Then LoginSession rsNew
            rsNew.Close
            Set rsNew = Nothing
            Flash "Bienvenido a PoliMusic, " & rUser & ".", "ok"
            Response.Redirect AppUrl("home.asp")
        Else
            rError = "No se pudo crear la cuenta. Intenta nuevamente."
        End If
    End If
End If

gTitle = "Crear cuenta"
%>
<!--#include file="includes/authheader.asp" -->

<h2>Crea tu cuenta</h2>
<p>Registrate para empezar a escuchar y armar tus playlists.</p>

<% If rError <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(rError)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("register.asp")%>" enctype="multipart/form-data" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-field">
        <label for="username">Usuario</label>
        <input type="text" id="username" name="username" class="pm-input" placeholder="tu_usuario"
               value="<%=H(rUser)%>" autocomplete="username" autofocus />
    </div>

    <div class="pm-field">
        <label for="email">Correo electronico</label>
        <input type="email" id="email" name="email" class="pm-input" placeholder="tucorreo@epn.edu.ec"
               value="<%=H(rEmail)%>" autocomplete="email" />
    </div>

    <div class="pm-field">
        <label for="birthday">Fecha de nacimiento</label>
        <input type="date" id="birthday" name="birthday" class="pm-input" value="<%=H(rBirth)%>" />
    </div>

    <div class="pm-field">
        <label for="password">Contrasena</label>
        <input type="password" id="password" name="password" class="pm-input"
               placeholder="Minimo 6 caracteres" autocomplete="new-password" />
    </div>

    <div class="pm-field">
        <label for="password2">Confirmar contrasena</label>
        <input type="password" id="password2" name="password2" class="pm-input"
               placeholder="Repite la contrasena" autocomplete="new-password" />
    </div>

    <div class="pm-field">
        <label for="photo">Foto de perfil (opcional)</label>
        <input type="file" id="photo" name="photo" class="pm-input" accept="image/*" />
        <span class="pm-hint">Si no subes una, se usara la imagen por defecto.</span>
    </div>

    <button type="submit" class="pm-btn pm-btn-primary" style="width:100%;padding:13px">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="12" cy="8" r="4" /><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8" />
        </svg>
        Crear cuenta
    </button>
</form>

<div class="pm-auth-foot">
    Ya tienes una cuenta? <a href="<%=AppUrl("login.asp")%>">Inicia sesion</a>
</div>

<!--#include file="includes/authfooter.asp" -->
