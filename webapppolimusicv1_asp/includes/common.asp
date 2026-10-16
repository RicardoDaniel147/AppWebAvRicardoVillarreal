<%
' =====================================================================
'  PoliMusic - Includes comunes (el orden es importante)
' =====================================================================
%>
<!--#include file="config.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="conn.asp" -->
<!--#include file="upload.asp" -->
<!--#include file="auth.asp" -->
<!--#include file="schema.asp" -->
<%
' Aplica los ajustes de esquema que la aplicacion necesita (una vez por arranque).
EnsureSchema()
%>
