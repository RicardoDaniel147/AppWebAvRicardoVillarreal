<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Listado de albumes
' =====================================================================
RequireLogin

Dim aSearch, aGenre, aCond, aGenres, rsAlbums, ai

aSearch = Qry("q")
aGenre  = Qry("genre")

aCond = ""
If aSearch <> "" Then aCond = aCond & IIf(aCond = "", "", " AND ") & "a.TITLE_ALBUM LIKE '%' + " & SqlText(aSearch) & " + '%'"
If aGenre <> ""  Then aCond = aCond & IIf(aCond = "", "", " AND ") & "a.GENRE = " & SqlText(aGenre)

Set rsAlbums = Query(SqlAlbumes(aCond, "a.TITLE_ALBUM", 0))
aGenres = GenerosAlbumes()

gPage = "albums"
gTitle = "Albumes"
%>
<!--#include file="includes/header.asp" -->

<h1 class="pm-h1">Albumes</h1>
<p class="pm-sub">Explora los albumes del catalogo.</p>

<div class="pm-panel" style="margin-bottom:22px">
    <form method="get" action="<%=AppUrl("albums.asp")%>" class="pm-form-grid">
        <div class="pm-field pm-span-2">
            <label for="q">Buscar album</label>
            <input type="search" id="q" name="q" class="pm-input" value="<%=H(aSearch)%>"
                   placeholder="Titulo del album..." autocomplete="off" />
        </div>

        <div class="pm-field">
            <label for="genre">Genero</label>
            <select id="genre" name="genre" class="pm-select">
                <option value="">Todos los generos</option>
                <% If UBound(aGenres) >= 0 Then
                       For ai = 0 To UBound(aGenres) %>
                <option value="<%=H(aGenres(ai))%>"<%=SelIf(aGenres(ai), aGenre)%>><%=H(aGenres(ai))%></option>
                <%     Next
                   End If %>
            </select>
        </div>

        <div class="pm-btn-row" style="align-self:end">
            <button type="submit" class="pm-btn pm-btn-primary">Filtrar</button>
            <a class="pm-btn" href="<%=AppUrl("albums.asp")%>">Limpiar</a>
        </div>
    </form>
</div>

<% If rsAlbums.EOF Then %>
<div class="pm-empty"><strong>Sin albumes</strong>No hay albumes que coincidan con el filtro.</div>
<% Else %>
<div class="pm-grid">
    <% Do While Not rsAlbums.EOF %>
    <a class="pm-card" href="<%=AppUrl("album.asp")%>?id=<%=rsAlbums("ID_ALBUM")%>">
        <div class="pm-card-art">
            <img src="<%=H(MediaUrl("" & rsAlbums("ALBUM_IMAGE_PATH"), "album", "" & rsAlbums("TITLE_ALBUM")))%>" alt="" loading="lazy" />
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsAlbums("TITLE_ALBUM"))%></div>
            <div class="pm-card-meta">
                <%=IIf(IsNull(rsAlbums("RELEASE_YEAR")), "", rsAlbums("RELEASE_YEAR") & " &middot; ")%><%=rsAlbums("CANCIONES")%> canciones
            </div>
        </div>
    </a>
    <% rsAlbums.MoveNext
       Loop %>
</div>
<% End If %>

<% rsAlbums.Close : Set rsAlbums = Nothing %>

<!--#include file="includes/footer.asp" -->
