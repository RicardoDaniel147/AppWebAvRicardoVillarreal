<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Cerrar sesion
' =====================================================================
LogoutSession()
Response.Redirect AppUrl("login.asp")
%>
