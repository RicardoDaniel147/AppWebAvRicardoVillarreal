<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Listado de artistas
' =====================================================================
RequireLogin

Dim arSearch, arCond, rsArtists

arSearch = Qry("q")
arCond = ""
If arSearch <> "" Then
    arCond = "a.ARTIST_NAME LIKE '%' + " & SqlText(arSearch) & " + '%' OR a.ARTIST_LASTNAME LIKE '%' + " & SqlText(arSearch) & " + '%'"
End If

Set rsArtists = Query(SqlArtistas(arCond, "a.ARTIST_NAME", 0))

gPage = "artists"
gTitle = "Artistas"
%>
<!--#include file="includes/header.asp" -->

<h1 class="pm-h1">Artistas</h1>
<p class="pm-sub">Todos los artistas registrados en el catalogo.</p>

<div class="pm-panel" style="margin-bottom:22px">
    <form method="get" action="<%=AppUrl("artists.asp")%>" class="pm-form-grid">
        <div class="pm-field pm-span-2">
            <label for="q">Buscar artista</label>
            <input type="search" id="q" name="q" class="pm-input" value="<%=H(arSearch)%>"
                   placeholder="Nombre o apellido..." autocomplete="off" />
        </div>
        <div class="pm-btn-row" style="align-self:end">
            <button type="submit" class="pm-btn pm-btn-primary">Buscar</button>
            <a class="pm-btn" href="<%=AppUrl("artists.asp")%>">Limpiar</a>
        </div>
    </form>
</div>

<% If rsArtists.EOF Then %>
<div class="pm-empty"><strong>Sin artistas</strong>No hay artistas que coincidan con la busqueda.</div>
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

<% rsArtists.Close : Set rsArtists = Nothing %>

<!--#include file="includes/footer.asp" -->
