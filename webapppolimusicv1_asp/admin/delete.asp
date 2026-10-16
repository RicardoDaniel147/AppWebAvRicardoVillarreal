<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Borrado generico (solo administradores)
'  Recibe POST: entity = user | song | album | artist | playlist | catalog
'               id     = identificador
'  Se encarga de borrar antes las filas relacionadas para respetar las FK.
' =====================================================================
RequireAdmin

Dim deEntity, deId, deBack, deName

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then
    Response.Redirect AppUrl("admin/dashboard.asp")
    Response.End
End If

ParseMultipart
CsrfCheck()

deEntity = LCase(PostVal("entity"))
deId     = ToInt(PostVal("id"), 0)
deBack   = AppUrl("admin/dashboard.asp")
deName   = ""

If deId <= 0 Then
    Flash "Identificador no valido.", "err"
    Response.Redirect deBack
    Response.End
End If

Select Case deEntity

    Case "user"
        deBack = AppUrl("admin/users.asp")
        deName = "" & Scalar("SELECT USERNAME FROM TBL_USER WHERE ID_USER = " & deId)

        If deId = CurrentUserId() Then
            Flash "No puedes eliminar tu propia cuenta mientras estas conectado.", "err"
            Response.Redirect deBack
            Response.End
        End If

        If deName <> "" Then
            Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = " & deId & ")"
            Execute "DELETE FROM TBL_PLAYLIST WHERE ID_USER = " & deId
            Execute "DELETE FROM TBL_USER WHERE ID_USER = " & deId
            Flash "Usuario """ & deName & """ eliminado (junto con sus playlists).", "ok"
        End If

    Case "song"
        deBack = AppUrl("admin/songs.asp")
        deName = "" & Scalar("SELECT SONG_NAME FROM TBL_SONG WHERE ID_SONG = " & deId)
        If deName <> "" Then
            Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = " & deId
            Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & deId
            Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & deId
            Execute "DELETE FROM TBL_SONG WHERE ID_SONG = " & deId
            Flash "Cancion """ & deName & """ eliminada.", "ok"
        End If

    Case "album"
        deBack = AppUrl("admin/albums.asp")
        deName = "" & Scalar("SELECT TITLE_ALBUM FROM TBL_ALBUM WHERE ID_ALBUM = " & deId)
        If deName <> "" Then
            Execute "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = " & deId
            Execute "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = " & deId
            Flash "Album """ & deName & """ eliminado.", "ok"
        End If

    Case "artist"
        deBack = AppUrl("admin/artists.asp")
        deName = "" & Scalar("SELECT ARTIST_NAME FROM TBL_ARTIST WHERE ID_ARTIST = " & deId)
        If deName <> "" Then
            Execute "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = " & deId
            Execute "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = " & deId
            Flash "Artista """ & deName & """ eliminado.", "ok"
        End If

    Case "playlist"
        deBack = AppUrl("admin/playlists.asp")
        deName = "" & Scalar("SELECT TITLE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & deId)
        If deName <> "" Then
            Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & deId
            Execute "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & deId
            Flash "Playlist """ & deName & """ eliminada.", "ok"
        End If

    Case "catalog"
        deBack = AppUrl("admin/catalog.asp")
        deName = "" & Scalar("SELECT VALUE FROM TBL_CATALOG WHERE ID_CATALOG = " & deId)
        If deName <> "" Then
            Execute "DELETE FROM TBL_CATALOG WHERE ID_CATALOG = " & deId
            Flash "Entrada de catalogo """ & deName & """ eliminada.", "ok"
        End If

    Case Else
        Flash "Entidad no reconocida.", "err"

End Select

Response.Redirect deBack
Response.End
%>
