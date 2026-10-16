<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Explorar: filtros combinables y pestanas de resultados
' =====================================================================
RequireLogin

Dim bSearch, bGenre, bArtist, bAlbum, bTab, bOrder, bCond, bAlbCond, bOrderSql
Dim rsSongs, rsAlbums, rsArtists, rsArtOpts, rsAlbOpts, bGenres, bi, nSongs, nAlbums, nArtists

bSearch = Qry("q")
bGenre  = Qry("genre")
bArtist = ToInt(Qry("artistId"), 0)
bAlbum  = ToInt(Qry("albumId"), 0)
bTab    = LCase(Qry("tab"))
bOrder  = LCase(Qry("order"))
If bTab = "" Then bTab = "songs"
If bOrder = "" Then bOrder = "name"

Select Case bOrder
    Case "plays" : bOrderSql = "s.PLAYS DESC, s.SONG_NAME"
    Case "genre" : bOrderSql = "s.GENRE, s.SONG_NAME"
    Case Else    : bOrderSql = "s.SONG_NAME"
End Select

bCond = ""
If bSearch <> "" Then bCond = bCond & IIf(bCond = "", "", " AND ") & "s.SONG_NAME LIKE '%' + " & SqlText(bSearch) & " + '%'"
If bGenre <> ""  Then bCond = bCond & IIf(bCond = "", "", " AND ") & "s.GENRE = " & SqlText(bGenre)
If bArtist > 0   Then bCond = bCond & IIf(bCond = "", "", " AND ") & "EXISTS (SELECT 1 FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG AND p.ID_ARTIST = " & bArtist & ")"
If bAlbum > 0    Then bCond = bCond & IIf(bCond = "", "", " AND ") & "EXISTS (SELECT 1 FROM TBL_TRACKLIST t WHERE t.ID_SONG = s.ID_SONG AND t.ID_ALBUM = " & bAlbum & ")"

bAlbCond = ""
If bSearch <> "" Then bAlbCond = bAlbCond & IIf(bAlbCond = "", "", " AND ") & "a.TITLE_ALBUM LIKE '%' + " & SqlText(bSearch) & " + '%'"
If bGenre <> ""  Then bAlbCond = bAlbCond & IIf(bAlbCond = "", "", " AND ") & "a.GENRE = " & SqlText(bGenre)

Set rsSongs    = Query(SqlCanciones(bCond, bOrderSql, 0))
Set rsAlbums   = Query(SqlAlbumes(bAlbCond, "a.TITLE_ALBUM", 0))
Set rsArtists  = Query(SqlArtistas(IIf(bSearch = "", "", "a.ARTIST_NAME LIKE '%' + " & SqlText(bSearch) & " + '%'"), "a.ARTIST_NAME", 0))
Set rsArtOpts  = Query(SqlArtistas("", "a.ARTIST_NAME", 0))
Set rsAlbOpts  = Query(SqlAlbumes("", "a.TITLE_ALBUM", 0))
bGenres = GenerosCanciones()

gPage = "search"
gTitle = "Explorar"
%>
<!--#include file="includes/header.asp" -->

<h1 class="pm-h1">Explorar</h1>
<p class="pm-sub">Filtra el catalogo por genero, artista y album, o busca por nombre.</p>

<div class="pm-panel" style="margin-bottom:26px">
    <form method="get" action="<%=AppUrl("search.asp")%>" class="pm-form-grid">
        <div class="pm-field pm-span-2">
            <label for="q">Buscar</label>
            <input type="search" id="q" name="q" class="pm-input" value="<%=H(bSearch)%>"
                   placeholder="Nombre de cancion, album o artista..." autocomplete="off" />
        </div>

        <div class="pm-field">
            <label for="genre">Genero</label>
            <select id="genre" name="genre" class="pm-select">
                <option value="">Todos los generos</option>
                <% If UBound(bGenres) >= 0 Then
                       For bi = 0 To UBound(bGenres) %>
                <option value="<%=H(bGenres(bi))%>"<%=SelIf(bGenres(bi), bGenre)%>><%=H(bGenres(bi))%></option>
                <%     Next
                   End If %>
            </select>
        </div>

        <div class="pm-field">
            <label for="artistId">Artista</label>
            <select id="artistId" name="artistId" class="pm-select">
                <option value="0">Todos los artistas</option>
                <% Do While Not rsArtOpts.EOF %>
                <option value="<%=rsArtOpts("ID_ARTIST")%>"<%=IIf(bArtist = ToInt(rsArtOpts("ID_ARTIST"), 0), " selected=""selected""", "")%>><%=H(NombreArtista(rsArtOpts("ARTIST_NAME"), rsArtOpts("ARTIST_LASTNAME")))%></option>
                <% rsArtOpts.MoveNext
                   Loop %>
            </select>
        </div>

        <div class="pm-field">
            <label for="albumId">Album</label>
            <select id="albumId" name="albumId" class="pm-select">
                <option value="0">Todos los albumes</option>
                <% Do While Not rsAlbOpts.EOF %>
                <option value="<%=rsAlbOpts("ID_ALBUM")%>"<%=IIf(bAlbum = ToInt(rsAlbOpts("ID_ALBUM"), 0), " selected=""selected""", "")%>><%=H(rsAlbOpts("TITLE_ALBUM"))%></option>
                <% rsAlbOpts.MoveNext
                   Loop %>
            </select>
        </div>

        <div class="pm-field">
            <label for="order">Ordenar por</label>
            <select id="order" name="order" class="pm-select">
                <option value="name"<%=SelIf("name", bOrder)%>>Nombre</option>
                <option value="plays"<%=SelIf("plays", bOrder)%>>Mas reproducidas</option>
                <option value="genre"<%=SelIf("genre", bOrder)%>>Genero</option>
            </select>
        </div>

        <input type="hidden" name="tab" value="<%=H(bTab)%>" />

        <div class="pm-span-2 pm-btn-row">
            <button type="submit" class="pm-btn pm-btn-primary">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <circle cx="11" cy="11" r="7" /><path d="m20 20-3.2-3.2" />
                </svg>
                Aplicar filtros
            </button>
            <a class="pm-btn" href="<%=AppUrl("search.asp")%>">Limpiar</a>
        </div>
    </form>
</div>

<% rsArtOpts.Close : Set rsArtOpts = Nothing %>
<% rsAlbOpts.Close : Set rsAlbOpts = Nothing %>

<%
' Enlaces de las pestanas conservando los filtros actuales
Dim qsBase
qsBase = "&q=" & Server.URLEncode(bSearch) & "&genre=" & Server.URLEncode(bGenre) & _
         "&artistId=" & bArtist & "&albumId=" & bAlbum & "&order=" & Server.URLEncode(bOrder)
%>

<div class="pm-admin-tabs">
    <a class="<%=IIf(bTab = "songs", "active", "")%>" href="<%=AppUrl("search.asp")%>?tab=songs<%=qsBase%>">Canciones</a>
    <a class="<%=IIf(bTab = "albums", "active", "")%>" href="<%=AppUrl("search.asp")%>?tab=albums<%=qsBase%>">Albumes</a>
    <a class="<%=IIf(bTab = "artists", "active", "")%>" href="<%=AppUrl("search.asp")%>?tab=artists<%=qsBase%>">Artistas</a>
</div>

<% If bTab = "albums" Then %>
    <% If rsAlbums.EOF Then %>
    <div class="pm-empty"><strong>Sin resultados</strong>No encontramos albumes con esos criterios.</div>
    <% Else %>
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

<% ElseIf bTab = "artists" Then %>
    <% If rsArtists.EOF Then %>
    <div class="pm-empty"><strong>Sin resultados</strong>No encontramos artistas con esos criterios.</div>
    <% Else %>
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

<% Else %>
    <!--#include file="includes/songlist.asp" -->
<% End If %>

<% rsSongs.Close : Set rsSongs = Nothing %>
<% rsAlbums.Close : Set rsAlbums = Nothing %>
<% rsArtists.Close : Set rsArtists = Nothing %>

<!--#include file="includes/footer.asp" -->
