<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Pagina de error
' =====================================================================
Dim eCode, eMsg, eIcono

eCode = ToInt(Qry("code"), 500)

Select Case eCode
    Case 400 : eMsg = "La solicitud no es valida." : eIcono = "&#9888;"
    Case 401 : eMsg = "Necesitas iniciar sesion." : eIcono = "&#128274;"
    Case 403 : eMsg = "No tienes permisos para ver este recurso." : eIcono = "&#128272;"
    Case 404 : eMsg = "No encontramos lo que buscabas." : eIcono = "&#128269;"
    Case 500 : eMsg = "Ocurrio un error inesperado en el servidor." : eIcono = "&#9888;"
    Case Else : eMsg = "Algo no salio como esperabamos." : eIcono = "&#9888;"
End Select

gTitle = "Error " & eCode
%>
<!--#include file="includes/authheader.asp" -->

<div style="text-align:center">
    <div style="font-size:54px;line-height:1;margin-bottom:12px"><%=eIcono%></div>
    <div class="pm-eyebrow">Error <%=eCode%></div>
    <h2 style="margin-top:10px"><%=H(eMsg)%></h2>

    <% If eCode = 403 And IsLoggedIn() And Not IsAdmin() Then %>
    <p>
        Esta seccion esta reservada a usuarios con rol <strong>administrador</strong>.
        Tu cuenta (<strong><%=H(CurrentUsername())%></strong>) no tiene los permisos necesarios.
    </p>
    <% Else %>
    <p>Intenta nuevamente o vuelve al inicio de la aplicacion.</p>
    <% End If %>

    <div class="pm-btn-row" style="justify-content:center;margin-top:26px">
        <% If IsLoggedIn() Then %>
        <a class="pm-btn pm-btn-primary" href="<%=AppUrl("home.asp")%>">Volver al inicio</a>
        <a class="pm-btn" href="<%=AppUrl("logout.asp")%>">Cambiar de cuenta</a>
        <% Else %>
        <a class="pm-btn pm-btn-primary" href="<%=AppUrl("login.asp")%>">Iniciar sesion</a>
        <% End If %>
    </div>
</div>

<!--#include file="includes/authfooter.asp" -->
