<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Acciones sobre una playlist: agregar, quitar o eliminar
'  Recibe POST desde playlist.asp y desde el listado de canciones.
' =====================================================================
RequireLogin

Dim iaAction, iaId, iaSong, rsPl, iaOwner, iaType

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
    Response.Redirect AppUrl("playlists.asp")
    Response.End
End If

ParseMultipart
CsrfCheck()

iaAction = LCase(PostVal("action"))
iaId     = ToInt(PostVal("id"), 0)
iaSong   = ToInt(PostVal("song"), 0)

If iaId <= 0 Then
    Response.Redirect AppUrl("playlists.asp")
    Response.End
End If

Set rsPl = Query("SELECT ID_PLAYLIST, ID_USER, TYPE, TITLE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & iaId)
If rsPl.EOF Then
    rsPl.Close : Set rsPl = Nothing
    Response.Redirect AppUrl("error.asp?code=404")
    Response.End
End If

iaOwner = ToInt(rsPl("ID_USER"), 0)
iaType  = ToInt(rsPl("TYPE"), PLAYLIST_PUBLIC)
rsPl.Close : Set rsPl = Nothing

If Not PuedeEditarPlaylist(iaId, iaOwner, iaType) Then
    Response.Redirect AppUrl("error.asp?code=403")
    Response.End
End If

Select Case iaAction
    Case "add"
        If iaSong > 0 Then
            Execute "IF NOT EXISTS (SELECT 1 FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & iaId & _
                    " AND ID_SONG = " & iaSong & ") " & _
                    "INSERT INTO TBL_PLAYLIST_SONG(ID_PLAYLIST, ID_SONG) VALUES(" & iaId & ", " & iaSong & ")"
            Flash "Cancion agregada a la playlist.", "ok"
        End If
        Response.Redirect AppUrl("playlist.asp") & "?id=" & iaId

    Case "remove"
        If iaSong > 0 Then
            Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & iaId & " AND ID_SONG = " & iaSong
            Flash "Cancion quitada de la playlist.", "ok"
        End If
        Response.Redirect AppUrl("playlist.asp") & "?id=" & iaId

    Case "delete"
        Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & iaId
        Execute "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & iaId
        Flash "Playlist eliminada correctamente.", "ok"
        Response.Redirect AppUrl("playlists.asp")

    Case Else
        Response.Redirect AppUrl("playlists.asp")
End Select

Response.End
%>
