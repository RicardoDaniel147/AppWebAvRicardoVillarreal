<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Administracion de artistas
' =====================================================================
RequireAdmin

Dim aaSearch, aaCond, rsArtists

aaSearch = Qry("q")
aaCond = ""
If aaSearch <> "" Then
    aaCond = "a.ARTIST_NAME LIKE '%' + " & SqlText(aaSearch) & " + '%' OR a.ARTIST_LASTNAME LIKE '%' + " & SqlText(aaSearch) & " + '%'"
End If

Set rsArtists = Query(SqlArtistas(aaCond, "a.ARTIST_NAME", 0))

gPage = "admin-artists"
gTitle = "Administrar artistas"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Artistas</h1>
<p class="pm-sub">Catalogo de artistas registrados.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-section" style="margin-top:0">
    <form method="get" action="<%=AppUrl("admin/artists.asp")%>" style="display:flex;gap:10px;align-items:center">
        <input type="search" name="q" class="pm-input" style="width:280px" value="<%=H(aaSearch)%>" placeholder="Buscar artista..." />
        <button type="submit" class="pm-btn pm-btn-sm">Buscar</button>
        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/artists.asp")%>">Limpiar</a>
    </form>

    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/artist_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nuevo artista
    </a>
</div>

<% If rsArtists.EOF Then %>
<div class="pm-empty"><strong>Sin artistas</strong>No hay artistas que coincidan con la busqueda.</div>
<% Else %>
<div class="pm-table-wrap">
    <table class="pm-table">
        <thead>
            <tr><th>Artista</th><th>Apellido</th><th>Canciones</th><th style="text-align:right">Acciones</th></tr>
        </thead>
        <tbody>
            <% Do While Not rsArtists.EOF %>
            <tr>
                <td>
                    <span class="cell-main">
                        <img class="round" src="<%=H(MediaUrl("" & rsArtists("ARTIST_IMAGE"), "artist", NombreArtista(rsArtists("ARTIST_NAME"), rsArtists("ARTIST_LASTNAME"))))%>" alt="" />
                        <strong><%=H(rsArtists("ARTIST_NAME"))%></strong>
                    </span>
                </td>
                <td><%=IIf(Trim("" & rsArtists("ARTIST_LASTNAME")) = "", "-", H("" & rsArtists("ARTIST_LASTNAME")))%></td>
                <td><%=rsArtists("CANCIONES")%></td>
                <td>
                    <span class="actions">
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/artist_edit.asp")%>?id=<%=rsArtists("ID_ARTIST")%>">Editar</a>
                        <form action="<%=AppUrl("admin/delete.asp")%>" method="post" style="margin:0"
                              onsubmit="return confirm('Eliminar el artista <%=JS(rsArtists("ARTIST_NAME"))%>? Se quitaran sus asociaciones con canciones.');">
                            <%=CsrfFieldInline()%>
                            <input type="hidden" name="entity" value="artist" />
                            <input type="hidden" name="id" value="<%=rsArtists("ID_ARTIST")%>" />
                            <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger">Eliminar</button>
                        </form>
                    </span>
                </td>
            </tr>
            <% rsArtists.MoveNext
               Loop %>
        </tbody>
    </table>
</div>
<% End If %>

<% rsArtists.Close : Set rsArtists = Nothing %>

<!--#include file="../includes/footer.asp" -->
