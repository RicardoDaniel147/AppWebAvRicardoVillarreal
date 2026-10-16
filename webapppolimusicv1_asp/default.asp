<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Punto de entrada
'  Redirige a la pantalla principal o al login segun la sesion.
' =====================================================================
If IsLoggedIn() Then
    Response.Redirect AppUrl("home.asp")
Else
    Response.Redirect AppUrl("login.asp")
End If
%>
