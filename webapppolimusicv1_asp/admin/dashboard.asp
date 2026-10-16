<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Panel con las metricas del sistema
' =====================================================================
RequireAdmin

Dim dbUsers, dbAdmins, dbSongs, dbArtists, dbAlbums, dbPlaylists, dbPlays
Dim rsTop, dbMax, dbPct, rsGenres, gTotal

dbUsers     = ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER"), 0)
dbAdmins    = ToInt(Scalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1"), 0)
dbSongs     = ToInt(Scalar("SELECT COUNT(*) FROM TBL_SONG"), 0)
dbArtists   = ToInt(Scalar("SELECT COUNT(*) FROM TBL_ARTIST"), 0)
dbAlbums    = ToInt(Scalar("SELECT COUNT(*) FROM TBL_ALBUM"), 0)
dbPlaylists = ToInt(Scalar("SELECT COUNT(*) FROM TBL_PLAYLIST"), 0)
dbPlays     = ToInt(Scalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG"), 0)

Set rsTop = Query(SqlCanciones("", "s.PLAYS DESC, s.SONG_NAME", 5))
dbMax = ToInt(Scalar("SELECT ISNULL(MAX(PLAYS),0) FROM TBL_SONG"), 0)
If dbMax <= 0 Then dbMax = 1

Set rsGenres = Query("SELECT ISNULL(GENRE,'Sin genero') AS G, COUNT(*) AS N FROM TBL_SONG GROUP BY GENRE ORDER BY COUNT(*) DESC")
gTotal = dbSongs
If gTotal <= 0 Then gTotal = 1

gPage = "admin-dashboard"
gTitle = "Panel de administracion"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Panel de administracion</h1>
<p class="pm-sub">Estado general de PoliMusic y acceso rapido a la gestion del catalogo.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-stats">
    <div class="pm-stat"><div class="k">Canciones</div><div class="v accent"><%=Num(dbSongs)%></div></div>
    <div class="pm-stat"><div class="k">Albumes</div><div class="v"><%=Num(dbAlbums)%></div></div>
    <div class="pm-stat"><div class="k">Artistas</div><div class="v"><%=Num(dbArtists)%></div></div>
    <div class="pm-stat"><div class="k">Playlists</div><div class="v"><%=Num(dbPlaylists)%></div></div>
    <div class="pm-stat"><div class="k">Usuarios</div><div class="v"><%=Num(dbUsers)%></div></div>
    <div class="pm-stat"><div class="k">Administradores</div><div class="v"><%=Num(dbAdmins)%></div></div>
    <div class="pm-stat"><div class="k">Reproducciones</div><div class="v"><%=Num(dbPlays)%></div></div>
</div>

<div class="pm-two-col">
    <div class="pm-panel">
        <h3>Top 5 canciones mas reproducidas</h3>
        <% If rsTop.EOF Then %>
        <div class="pm-empty">Sin datos de reproduccion.</div>
        <% Else %>
        <div class="pm-bar-chart">
            <% Do While Not rsTop.EOF
                   dbPct = ToInt(rsTop("PLAYS"), 0) * 100 / dbMax %>
            <div class="pm-bar-item">
                <span class="g" title="<%=H(rsTop("SONG_NAME"))%>"><%=H(rsTop("SONG_NAME"))%></span>
                <span class="track"><span class="fill" style="width:<%=FormatNumber(dbPct, 1, -1, 0, 0)%>%"></span></span>
                <span class="n"><%=Num(rsTop("PLAYS"))%></span>
            </div>
            <% rsTop.MoveNext
               Loop %>
        </div>
        <% End If %>
    </div>

    <div class="pm-panel">
        <h3>Distribucion por genero</h3>
        <% If rsGenres.EOF Then %>
        <div class="pm-empty">Sin generos registrados.</div>
        <% Else %>
        <div class="pm-bar-chart">
            <% Do While Not rsGenres.EOF
                   dbPct = ToInt(rsGenres("N"), 0) * 100 / gTotal %>
            <div class="pm-bar-item">
                <span class="g" title="<%=H(rsGenres("G"))%>"><%=H(rsGenres("G"))%></span>
                <span class="track"><span class="fill" style="width:<%=FormatNumber(dbPct, 1, -1, 0, 0)%>%"></span></span>
                <span class="n"><%=rsGenres("N")%></span>
            </div>
            <% rsGenres.MoveNext
               Loop %>
        </div>
        <% End If %>
    </div>
</div>

<div class="pm-section">
    <h2>Accesos rapidos</h2>
</div>

<div class="pm-btn-row">
    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/song_edit.asp")%>">Nueva cancion</a>
    <a class="pm-btn" href="<%=AppUrl("admin/album_edit.asp")%>">Nuevo album</a>
    <a class="pm-btn" href="<%=AppUrl("admin/artist_edit.asp")%>">Nuevo artista</a>
    <a class="pm-btn" href="<%=AppUrl("admin/user_edit.asp")%>">Nuevo usuario</a>
    <a class="pm-btn" href="<%=AppUrl("admin/playlist_edit.asp")%>">Nueva playlist</a>
    <a class="pm-btn" href="<%=AppUrl("admin/catalog.asp")%>">Catalogo</a>
</div>

<% rsTop.Close : Set rsTop = Nothing %>
<% rsGenres.Close : Set rsGenres = Nothing %>

<!--#include file="../includes/footer.asp" -->
