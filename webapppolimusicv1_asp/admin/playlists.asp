<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Administracion de playlists de todos los usuarios
' =====================================================================
RequireAdmin

Dim apSearch, apCond, rsPlaylists

apSearch = Qry("q")
apCond = ""
If apSearch <> "" Then
    apCond = "pl.TITLE LIKE '%' + " & SqlText(apSearch) & " + '%' OR u.USERNAME LIKE '%' + " & SqlText(apSearch) & " + '%'"
End If

Set rsPlaylists = Query(SqlPlaylists(apCond, "pl.CREATION_DATE DESC", 0))

gPage = "admin-playlists"
gTitle = "Administrar playlists"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Playlists</h1>
<p class="pm-sub">Todas las playlists del sistema, de cualquier usuario.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-section" style="margin-top:0">
    <form method="get" action="<%=AppUrl("admin/playlists.asp")%>" style="display:flex;gap:10px;align-items:center">
        <input type="search" name="q" class="pm-input" style="width:280px" value="<%=H(apSearch)%>"
               placeholder="Buscar por titulo o propietario..." />
        <button type="submit" class="pm-btn pm-btn-sm">Buscar</button>
        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/playlists.asp")%>">Limpiar</a>
    </form>

    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/playlist_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nueva playlist
    </a>
</div>

<% If rsPlaylists.EOF Then %>
<div class="pm-empty"><strong>Sin playlists</strong>No hay playlists que coincidan con la busqueda.</div>
<% Else %>
<div class="pm-table-wrap">
    <table class="pm-table">
        <thead>
            <tr>
                <th>Playlist</th><th>Propietario</th><th>Tipo</th>
                <th>Canciones</th><th>Creada</th><th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Do While Not rsPlaylists.EOF %>
            <tr>
                <td>
                    <span class="cell-main">
                        <img src="<%=H(MediaUrl("" & rsPlaylists("PHOTO"), "playlist", "" & rsPlaylists("TITLE")))%>" alt="" />
                        <strong><%=H(rsPlaylists("TITLE"))%></strong>
                    </span>
                </td>
                <td><%=IIf(IsNull(rsPlaylists("PROPIETARIO")), "-", H("" & rsPlaylists("PROPIETARIO")))%></td>
                <td><span class="<%=BadgePlaylist(rsPlaylists("TYPE"))%>"><%=TipoPlaylistNombre(rsPlaylists("TYPE"))%></span></td>
                <td><%=rsPlaylists("CANCIONES")%></td>
                <td><%=FechaCorta(rsPlaylists("CREATION_DATE"))%></td>
                <td>
                    <span class="actions">
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("playlist.asp")%>?id=<%=rsPlaylists("ID_PLAYLIST")%>">Ver</a>
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/playlist_edit.asp")%>?id=<%=rsPlaylists("ID_PLAYLIST")%>">Editar</a>
                        <form action="<%=AppUrl("admin/delete.asp")%>" method="post" style="margin:0"
                              onsubmit="return confirm('Eliminar la playlist <%=JS(rsPlaylists("TITLE"))%>?');">
                            <%=CsrfFieldInline()%>
                            <input type="hidden" name="entity" value="playlist" />
                            <input type="hidden" name="id" value="<%=rsPlaylists("ID_PLAYLIST")%>" />
                            <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger">Eliminar</button>
                        </form>
                    </span>
                </td>
            </tr>
            <% rsPlaylists.MoveNext
               Loop %>
        </tbody>
    </table>
</div>
<% End If %>

<% rsPlaylists.Close : Set rsPlaylists = Nothing %>

<!--#include file="../includes/footer.asp" -->
