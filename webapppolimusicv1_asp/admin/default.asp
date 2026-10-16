<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Entrada al panel de administracion
' =====================================================================
RequireAdmin
Response.Redirect AppUrl("admin/dashboard.asp")
Response.End
%>
