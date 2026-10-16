<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Detalle de un album con su tracklist
' =====================================================================
RequireLogin

Dim alId, rsAlbum, rsSongs, alTotal

alId = ToInt(Qry("id"), 0)
Set rsAlbum = Query(SqlAlbumes("a.ID_ALBUM = " & alId, "", 0))
If rsAlbum.EOF Then
    rsAlbum.Close
    Set rsAlbum = Nothing
    Response.Redirect AppUrl("error.asp?code=404")
    Response.End
End If

Set rsSongs = Query(SqlCanciones("EXISTS (SELECT 1 FROM TBL_TRACKLIST t WHERE t.ID_SONG = s.ID_SONG AND t.ID_ALBUM = " & alId & ")", "s.SONG_NAME", 0))
alTotal = ToInt(Scalar("SELECT ISNULL(SUM(s.PLAYS),0) FROM TBL_TRACKLIST t INNER JOIN TBL_SONG s ON s.ID_SONG = t.ID_SONG WHERE t.ID_ALBUM = " & alId), 0)

gPage = "albums"
gTitle = "" & rsAlbum("TITLE_ALBUM")
%>
<!--#include file="includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("albums.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a albumes
</a>

<section class="pm-detail pm-fade-in">
    <div class="pm-detail-art">
        <img src="<%=H(MediaUrl("" & rsAlbum("ALBUM_IMAGE_PATH"), "album", "" & rsAlbum("TITLE_ALBUM")))%>"
             alt="<%=H(rsAlbum("TITLE_ALBUM"))%>" />
    </div>

    <div>
        <span class="pm-eyebrow">Album</span>
        <h1><%=H(rsAlbum("TITLE_ALBUM"))%></h1>
        <div class="pm-detail-meta">
            <span class="pm-badge pm-badge-user"><%=H("" & rsAlbum("GENRE"))%></span>
            <% If Not IsNull(rsAlbum("RELEASE_YEAR")) Then %>
            <span><%=rsAlbum("RELEASE_YEAR")%></span>
            <span class="dot"></span>
            <% End If %>
            <span><%=rsAlbum("CANCIONES")%> canciones</span>
            <span class="dot"></span>
            <span><%=Num(alTotal)%> reproducciones</span>
        </div>

        <div class="pm-btn-row" style="margin-top:22px">
            <button type="button" class="pm-btn pm-btn-primary" onclick="PoliMusic.playFirst()">
                <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z" /></svg>
                Reproducir album
            </button>
            <a class="pm-btn" href="<%=AppUrl("search.asp")%>?genre=<%=Server.URLEncode("" & rsAlbum("GENRE"))%>">
                Mas de <%=H("" & rsAlbum("GENRE"))%>
            </a>
        </div>
    </div>
</section>

<% rsAlbum.Close : Set rsAlbum = Nothing %>

<div class="pm-section">
    <h2>Lista de canciones</h2>
</div>

<!--#include file="includes/songlist.asp" -->

<% rsSongs.Close : Set rsSongs = Nothing %>

<!--#include file="includes/footer.asp" -->
