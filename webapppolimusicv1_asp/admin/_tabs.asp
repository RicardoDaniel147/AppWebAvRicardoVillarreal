<%
' =====================================================================
'  PoliMusic - Pestanas del panel de administracion
'  Requiere que la pagina haya asignado gPage.
' =====================================================================
%>
<div class="pm-admin-tabs">
    <a class="<%=NavActivo("admin-dashboard")%>" href="<%=AppUrl("admin/dashboard.asp")%>">Panel</a>
    <a class="<%=NavActivo("admin-users")%>" href="<%=AppUrl("admin/users.asp")%>">Usuarios</a>
    <a class="<%=NavActivo("admin-songs")%>" href="<%=AppUrl("admin/songs.asp")%>">Canciones</a>
    <a class="<%=NavActivo("admin-albums")%>" href="<%=AppUrl("admin/albums.asp")%>">Albumes</a>
    <a class="<%=NavActivo("admin-artists")%>" href="<%=AppUrl("admin/artists.asp")%>">Artistas</a>
    <a class="<%=NavActivo("admin-playlists")%>" href="<%=AppUrl("admin/playlists.asp")%>">Playlists</a>
    <a class="<%=NavActivo("admin-catalog")%>" href="<%=AppUrl("admin/catalog.asp")%>">Catalogo</a>
</div>
