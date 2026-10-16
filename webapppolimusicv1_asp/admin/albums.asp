<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Administracion de albumes
' =====================================================================
RequireAdmin

Dim abSearch, abGenre, abCond, abGenres, rsAlbums, abi

abSearch = Qry("q")
abGenre  = Qry("genre")

abCond = ""
If abSearch <> "" Then abCond = abCond & IIf(abCond = "", "", " AND ") & "a.TITLE_ALBUM LIKE '%' + " & SqlText(abSearch) & " + '%'"
If abGenre <> ""  Then abCond = abCond & IIf(abCond = "", "", " AND ") & "a.GENRE = " & SqlText(abGenre)

Set rsAlbums = Query(SqlAlbumes(abCond, "a.TITLE_ALBUM", 0))
abGenres = GenerosAlbumes()

gPage = "admin-albums"
gTitle = "Administrar albumes"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Albumes</h1>
<p class="pm-sub">Gestiona los albumes y su tracklist.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-section" style="margin-top:0">
    <form method="get" action="<%=AppUrl("admin/albums.asp")%>" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <input type="search" name="q" class="pm-input" style="width:250px" value="<%=H(abSearch)%>" placeholder="Buscar album..." />
        <select name="genre" class="pm-select" style="width:180px">
            <option value="">Todos los generos</option>
            <% If UBound(abGenres) >= 0 Then
                   For abi = 0 To UBound(abGenres) %>
            <option value="<%=H(abGenres(abi))%>"<%=SelIf(abGenres(abi), abGenre)%>><%=H(abGenres(abi))%></option>
            <%     Next
               End If %>
        </select>
        <button type="submit" class="pm-btn pm-btn-sm">Filtrar</button>
        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/albums.asp")%>">Limpiar</a>
    </form>

    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/album_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nuevo album
    </a>
</div>

<% If rsAlbums.EOF Then %>
<div class="pm-empty"><strong>Sin albumes</strong>No hay albumes que coincidan con el filtro.</div>
<% Else %>
<div class="pm-table-wrap">
    <table class="pm-table">
        <thead>
            <tr><th>Album</th><th>Genero</th><th>Ano</th><th>Canciones</th><th style="text-align:right">Acciones</th></tr>
        </thead>
        <tbody>
            <% Do While Not rsAlbums.EOF %>
            <tr>
                <td>
                    <span class="cell-main">
                        <img src="<%=H(MediaUrl("" & rsAlbums("ALBUM_IMAGE_PATH"), "album", "" & rsAlbums("TITLE_ALBUM")))%>" alt="" />
                        <strong><%=H(rsAlbums("TITLE_ALBUM"))%></strong>
                    </span>
                </td>
                <td><span class="pm-badge pm-badge-user"><%=H("" & rsAlbums("GENRE"))%></span></td>
                <td><%=IIf(IsNull(rsAlbums("RELEASE_YEAR")), "-", rsAlbums("RELEASE_YEAR"))%></td>
                <td><%=rsAlbums("CANCIONES")%></td>
                <td>
                    <span class="actions">
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("album.asp")%>?id=<%=rsAlbums("ID_ALBUM")%>">Ver</a>
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/album_edit.asp")%>?id=<%=rsAlbums("ID_ALBUM")%>">Editar</a>
                        <form action="<%=AppUrl("admin/delete.asp")%>" method="post" style="margin:0"
                              onsubmit="return confirm('Eliminar el album <%=JS(rsAlbums("TITLE_ALBUM"))%>? Se eliminara su tracklist.');">
                            <%=CsrfFieldInline()%>
                            <input type="hidden" name="entity" value="album" />
                            <input type="hidden" name="id" value="<%=rsAlbums("ID_ALBUM")%>" />
                            <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger">Eliminar</button>
                        </form>
                    </span>
                </td>
            </tr>
            <% rsAlbums.MoveNext
               Loop %>
        </tbody>
    </table>
</div>
<% End If %>

<% rsAlbums.Close : Set rsAlbums = Nothing %>

<!--#include file="../includes/footer.asp" -->
