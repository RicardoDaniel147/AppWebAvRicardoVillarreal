<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Detalle de un artista con sus canciones
' =====================================================================
RequireLogin

Dim arId, rsArtist, rsSongs, arNombre, arTotal

arId = ToInt(Qry("id"), 0)
Set rsArtist = Query(SqlArtistas("a.ID_ARTIST = " & arId, "", 0))
If rsArtist.EOF Then
    rsArtist.Close
    Set rsArtist = Nothing
    Response.Redirect AppUrl("error.asp?code=404")
    Response.End
End If

arNombre = NombreArtista(rsArtist("ARTIST_NAME"), rsArtist("ARTIST_LASTNAME"))
Set rsSongs = Query(SqlCanciones("EXISTS (SELECT 1 FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG AND p.ID_ARTIST = " & arId & ")", "s.SONG_NAME", 0))
arTotal = ToInt(Scalar("SELECT ISNULL(SUM(s.PLAYS),0) FROM TBL_PLAYER p INNER JOIN TBL_SONG s ON s.ID_SONG = p.ID_SONG WHERE p.ID_ARTIST = " & arId), 0)

gPage = "artists"
gTitle = arNombre
%>
<!--#include file="includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("artists.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a artistas
</a>

<section class="pm-detail pm-fade-in">
    <div class="pm-detail-art round">
        <img src="<%=H(MediaUrl("" & rsArtist("ARTIST_IMAGE"), "artist", arNombre))%>" alt="<%=H(arNombre)%>" />
    </div>

    <div>
        <span class="pm-eyebrow">Artista</span>
        <h1><%=H(arNombre)%></h1>
        <div class="pm-detail-meta">
            <span><%=rsArtist("CANCIONES")%> canciones</span>
            <span class="dot"></span>
            <span><%=Num(arTotal)%> reproducciones</span>
        </div>

        <div class="pm-btn-row" style="margin-top:22px">
            <button type="button" class="pm-btn pm-btn-primary" onclick="PoliMusic.playFirst()">
                <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z" /></svg>
                Reproducir todo
            </button>
        </div>
    </div>
</section>

<% rsArtist.Close : Set rsArtist = Nothing %>

<div class="pm-section">
    <h2>Canciones de <%=H(arNombre)%></h2>
</div>

<!--#include file="includes/songlist.asp" -->

<% rsSongs.Close : Set rsSongs = Nothing %>

<!--#include file="includes/footer.asp" -->
