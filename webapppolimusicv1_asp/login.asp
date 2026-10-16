<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Iniciar sesion
' =====================================================================
Dim lUser, lPass, lError, lNext, rsLogin

lError = ""
lNext = Qry("next")
If lNext = "" Then lNext = AppUrl("home.asp")

If IsLoggedIn() Then Response.Redirect AppUrl("home.asp")

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    CsrfCheck()
    lUser = PostVal("username")
    lPass = PostVal("password")

    If lUser = "" Or lPass = "" Then
        lError = "Ingresa tu usuario y tu contrasena."
    Else
        ' Acepta contrasenas en texto plano (datos heredados) y hasheadas en SHA1.
        Set rsLogin = Query("SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER " & _
                            "WHERE USERNAME = " & SqlText(lUser) & _
                            " AND (PASSWORD = " & SqlText(lPass) & _
                            " OR UPPER(PASSWORD) = UPPER(CONVERT(VARCHAR(40), HASHBYTES('SHA1', " & SqlAscii(lPass) & "), 2)))")

        If rsLogin.EOF Then
            lError = "Usuario o contrasena incorrectos."
        Else
            LoginSession rsLogin
            rsLogin.Close
            Set rsLogin = Nothing

            If Left(lNext, 1) = "/" Then
                Response.Redirect lNext
            Else
                Response.Redirect AppUrl("home.asp")
            End If
        End If

        If IsObject(rsLogin) Then
            If rsLogin.State = 1 Then rsLogin.Close
            Set rsLogin = Nothing
        End If
    End If
End If

gTitle = "Iniciar sesion"
%>
<!--#include file="includes/authheader.asp" -->

<h2>Inicia sesion</h2>
<p>Accede a tu biblioteca musical en PoliMusic.</p>

<% If lError <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(lError)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("login.asp")%>?next=<%=Server.URLEncode(lNext)%>" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-field">
        <label for="username">Usuario</label>
        <input type="text" id="username" name="username" class="pm-input" placeholder="userAdmin"
               value="<%=H(lUser)%>" autocomplete="username" autofocus />
    </div>

    <div class="pm-field">
        <label for="password">Contrasena</label>
        <input type="password" id="password" name="password" class="pm-input" placeholder="********"
               autocomplete="current-password" />
    </div>

    <label class="pm-check">
        <input type="checkbox" name="remember" value="1" checked="checked" />
        <span>Mantener la sesion iniciada</span>
    </label>

    <button type="submit" class="pm-btn pm-btn-primary" style="width:100%;padding:13px">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4" /><path d="M10 17l5-5-5-5" /><path d="M15 12H3" />
        </svg>
        Entrar
    </button>
</form>

<div class="pm-divider">cuentas de prueba</div>

<div class="pm-demo">
    <strong style="color:var(--text-dim)">Administrador</strong> &nbsp;
    <code>userAdmin</code> / <code>12345678</code><br />
    <strong style="color:var(--text-dim)">Usuario normal</strong> &nbsp;
    <code>userNormal</code> / <code>12345678</code>
</div>

<div class="pm-auth-foot">
    No tienes cuenta? <a href="<%=AppUrl("register.asp")%>">Registrate aqui</a>
</div>

<!--#include file="includes/authfooter.asp" -->
