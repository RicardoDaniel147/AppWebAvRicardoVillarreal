<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Pantalla principal
' =====================================================================
RequireLogin

Dim hGenre, hGreeting, hh, hGenres, hCond, hAlbCond
Dim rsSongs, rsTop, rsAlbums, rsArtists, rsPlaylists

hGenre = Qry("genre")
hh = Hour(Now())
If hh < 12 Then
    hGreeting = "Buenos dias"
ElseIf hh < 19 Then
    hGreeting = "Buenas tardes"
Else
    hGreeting = "Buenas noches"
End If

hGenres = GenerosCanciones()
hCond = ""
hAlbCond = ""
If hGenre <> "" Then
    hCond = "s.GENRE = " & SqlText(hGenre)
    hAlbCond = "a.GENRE = " & SqlText(hGenre)
End If

Set rsSongs     = Query(SqlCanciones(hCond, "s.SONG_NAME", 12))
Set rsTop       = Query(SqlCanciones(hCond, "s.PLAYS DESC, s.SONG_NAME", 6))
Set rsAlbums    = Query(SqlAlbumes(hAlbCond, "a.TITLE_ALBUM", 8))
Set rsArtists   = Query(SqlArtistas("", "a.ARTIST_NAME", 8))
Set rsPlaylists = Query(SqlPlaylists("pl.TYPE = 0", "pl.CREATION_DATE DESC", 6))

gPage = "home"
gTitle = "Inicio"
%>
<!--#include file="includes/header.asp" -->

<section class="pm-hero pm-fade-in">
    <div class="pm-hero-content">
        <span class="pm-eyebrow">Repositorio musical</span>
        <h1><%=H(hGreeting)%>, <em><%=H(CurrentUsername())%></em>.<br />Tu musica esta lista.</h1>
        <p>
            <% If hGenre <> "" Then %>
            Filtrando por el genero <strong><%=H(hGenre)%></strong>.
            <% Else %>
            Explora el catalogo completo, arma tus playlists y reproduce sin interrupciones.
            <% End If %>
        </p>

        <div class="pm-btn-row">
            <button type="button" class="pm-btn pm-btn-primary" onclick="PoliMusic.playFirst()">
                <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z" /></svg>
                Reproducir todo
            </button>
            <a class="pm-btn" href="<%=AppUrl("search.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <circle cx="11" cy="11" r="7" /><path d="m20 20-3.2-3.2" />
                </svg>
                Explorar catalogo
            </a>
            <% If IsAdmin() Then %>
            <a class="pm-btn pm-btn-cyan" href="<%=AppUrl("admin/dashboard.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <path d="M4 13h6V4H4zM14 20h6v-9h-6zM4 20h6v-4H4zM14 8h6V4h-6z" />
                </svg>
                Panel de administracion
            </a>
            <% End If %>
        </div>
    </div>
</section>

<% If IsAdmin() Then %>
<div class="pm-section">
    <h2>Resumen del sistema</h2>
    <a href="<%=AppUrl("admin/dashboard.asp")%>">Ver detalle</a>
</div>

<div class="pm-stats">
    <div class="pm-stat"><div class="k">Canciones</div><div class="v accent"><%=Num(Scalar("SELECT COUNT(*) FROM TBL_SONG"))%></div></div>
    <div class="pm-stat"><div class="k">Albumes</div><div class="v"><%=Num(Scalar("SELECT COUNT(*) FROM TBL_ALBUM"))%></div></div>
    <div class="pm-stat"><div class="k">Artistas</div><div class="v"><%=Num(Scalar("SELECT COUNT(*) FROM TBL_ARTIST"))%></div></div>
    <div class="pm-stat"><div class="k">Playlists</div><div class="v"><%=Num(Scalar("SELECT COUNT(*) FROM TBL_PLAYLIST"))%></div></div>
    <div class="pm-stat"><div class="k">Usuarios</div><div class="v"><%=Num(Scalar("SELECT COUNT(*) FROM TBL_USER"))%></div></div>
    <div class="pm-stat"><div class="k">Reproducciones</div><div class="v"><%=Num(Scalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG"))%></div></div>
</div>
<% End If %>

<% If UBound(hGenres) >= 0 Then %>
<div class="pm-section">
    <h2>Filtrar por genero</h2>
    <% If hGenre <> "" Then %><a href="<%=AppUrl("home.asp")%>">Quitar filtro</a><% End If %>
</div>

<div class="pm-chips" style="margin-bottom:8px">
    <a class="pm-chip <%=IIf(hGenre = "", "active", "")%>" href="<%=AppUrl("home.asp")%>">Todos</a>
    <% Dim gi
       For gi = 0 To UBound(hGenres) %>
    <a class="pm-chip <%=IIf(LCase(hGenre) = LCase(hGenres(gi)), "active", "")%>"
       href="<%=AppUrl("home.asp")%>?genre=<%=Server.URLEncode(hGenres(gi))%>"><%=H(hGenres(gi))%></a>
    <% Next %>
</div>
<% End If %>

<div class="pm-section">
    <h2>Canciones</h2>
    <a href="<%=AppUrl("songs.asp")%>">Ver todas</a>
</div>

<!--#include file="includes/songlist.asp" -->

<% rsSongs.Close : Set rsSongs = Nothing %>

<% If Not rsTop.EOF Then %>
<div class="pm-section">
    <h2>Mas reproducidas</h2>
    <a href="<%=AppUrl("songs.asp")%>?order=plays">Ranking completo</a>
</div>

<div class="pm-grid wide">
    <% Dim topCover, topSeed
       Do While Not rsTop.EOF
           topSeed = "" & rsTop("ARTISTAS")
           topCover = MediaUrl("" & rsTop("ARTIST_IMAGE"), "artist", topSeed) %>
    <div class="pm-card">
        <div class="pm-card-art">
            <img src="<%=H(topCover)%>" alt="" loading="lazy" />
            <button type="button" class="pm-card-play" title="Reproducir" onclick="PoliMusic.play(<%=rsTop("ID_SONG")%>)">
                <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z" /></svg>
            </button>
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsTop("SONG_NAME"))%></div>
            <div class="pm-card-meta"><%=Num(rsTop("PLAYS"))%> reproducciones</div>
        </div>
    </div>
    <% rsTop.MoveNext
       Loop %>
</div>
<% End If %>
<% rsTop.Close : Set rsTop = Nothing %>

<% If Not rsAlbums.EOF Then %>
<div class="pm-section">
    <h2>Albumes</h2>
    <a href="<%=AppUrl("albums.asp")%>">Ver todos</a>
</div>

<div class="pm-grid">
    <% Do While Not rsAlbums.EOF %>
    <a class="pm-card" href="<%=AppUrl("album.asp")%>?id=<%=rsAlbums("ID_ALBUM")%>">
        <div class="pm-card-art">
            <img src="<%=H(MediaUrl("" & rsAlbums("ALBUM_IMAGE_PATH"), "album", "" & rsAlbums("TITLE_ALBUM")))%>" alt="" loading="lazy" />
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsAlbums("TITLE_ALBUM"))%></div>
            <div class="pm-card-meta"><%=rsAlbums("CANCIONES")%> canciones &middot; <%=H("" & rsAlbums("GENRE"))%></div>
        </div>
    </a>
    <% rsAlbums.MoveNext
       Loop %>
</div>
<% End If %>
<% rsAlbums.Close : Set rsAlbums = Nothing %>

<% If Not rsArtists.EOF Then %>
<div class="pm-section">
    <h2>Artistas</h2>
    <a href="<%=AppUrl("artists.asp")%>">Ver todos</a>
</div>

<div class="pm-grid artists">
    <% Do While Not rsArtists.EOF %>
    <a class="pm-card" href="<%=AppUrl("artist.asp")%>?id=<%=rsArtists("ID_ARTIST")%>">
        <div class="pm-card-art round">
            <img src="<%=H(MediaUrl("" & rsArtists("ARTIST_IMAGE"), "artist", NombreArtista(rsArtists("ARTIST_NAME"), rsArtists("ARTIST_LASTNAME"))))%>" alt="" loading="lazy" />
        </div>
        <div style="text-align:center">
            <div class="pm-card-title"><%=H(NombreArtista(rsArtists("ARTIST_NAME"), rsArtists("ARTIST_LASTNAME")))%></div>
            <div class="pm-card-meta"><%=rsArtists("CANCIONES")%> canciones</div>
        </div>
    </a>
    <% rsArtists.MoveNext
       Loop %>
</div>
<% End If %>
<% rsArtists.Close : Set rsArtists = Nothing %>

<% If Not rsPlaylists.EOF Then %>
<div class="pm-section">
    <h2>Playlists publicas</h2>
    <a href="<%=AppUrl("playlists.asp")%>">Mis playlists</a>
</div>

<div class="pm-grid">
    <% Do While Not rsPlaylists.EOF %>
    <a class="pm-card" href="<%=AppUrl("playlist.asp")%>?id=<%=rsPlaylists("ID_PLAYLIST")%>">
        <div class="pm-card-art">
            <img src="<%=H(MediaUrl("" & rsPlaylists("PHOTO"), "playlist", "" & rsPlaylists("TITLE")))%>" alt="" loading="lazy" />
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsPlaylists("TITLE"))%></div>
            <div class="pm-card-meta"><%=rsPlaylists("CANCIONES")%> canciones &middot; <%=H("" & rsPlaylists("PROPIETARIO"))%></div>
        </div>
    </a>
    <% rsPlaylists.MoveNext
       Loop %>
</div>
<% End If %>
<% rsPlaylists.Close : Set rsPlaylists = Nothing %>

<!--#include file="includes/footer.asp" -->
