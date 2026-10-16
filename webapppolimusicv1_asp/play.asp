<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Reproduccion
'   POST -> incrementa el contador de reproducciones y responde JSON
'   GET  -> incrementa y redirige al archivo de audio
' =====================================================================
RequireLogin

Dim pkId, pkSong, pkPath, pkUrl, pkPlays, pkToken, pkMethod

pkId     = ToInt(R("id"), 0)
pkMethod = UCase("" & Request.ServerVariables("REQUEST_METHOD"))

If pkId <= 0 Then
    Response.Status = "400 Bad Request"
    Response.ContentType = "application/json"
    Response.Write "{""ok"":false,""error"":""id requerido""}"
    Response.End
End If

' Proteccion CSRF: acepta el campo _t del formulario o la cabecera del reproductor.
If pkMethod = "POST" Then
    pkToken = "" & Request.Form("_t")
    If pkToken = "" Then pkToken = "" & Request.ServerVariables("HTTP_X_CSRF_TOKEN")
    If pkToken = "" Or pkToken <> ("" & Session("csrf")) Then
        Response.Status = "403 Forbidden"
        Response.ContentType = "application/json"
        Response.Write "{""ok"":false,""error"":""token invalido""}"
        Response.End
    End If
End If

pkSong = "" & Scalar("SELECT SONG_NAME FROM TBL_SONG WHERE ID_SONG = " & pkId)
If pkSong = "" Then
    Response.Status = "404 Not Found"
    Response.ContentType = "application/json"
    Response.Write "{""ok"":false,""error"":""cancion no encontrada""}"
    Response.End
End If

Execute "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS,0) + 1 WHERE ID_SONG = " & pkId
pkPlays = ToInt(Scalar("SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = " & pkId), 0)

If pkMethod = "POST" Then
    Response.ContentType = "application/json"
    Response.Write "{""ok"":true,""id"":" & pkId & ",""plays"":" & pkPlays & "}"
    Response.End
End If

' GET: redirige al archivo real para que el navegador pueda hacer "seeking".
pkPath = "" & Scalar("SELECT SONG_PATH FROM TBL_SONG WHERE ID_SONG = " & pkId)
pkUrl  = AudioUrl(pkPath)

If pkUrl = "" Then
    Response.Redirect AppUrl("error.asp?code=404")
    Response.End
End If

Response.Redirect pkUrl
Response.End
%>
