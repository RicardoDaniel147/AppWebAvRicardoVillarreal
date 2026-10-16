<%
' =====================================================================
'  PoliMusic - Cabecera comun (sidebar + topbar)
'  La pagina debe asignar gPage y gTitle antes de incluir este archivo.
' =====================================================================
Dim gNavQ
gNavQ = "" & Request.QueryString("q")
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title><%=H(gTitle)%> &middot; <%=H(SITE_NAME)%></title>
    <link rel="stylesheet" href="<%=Asset("style.css")%>" />
</head>
<body>
<div class="pm-bg"></div>

<div class="pm-app">

    <!-- ============================================================ SIDEBAR -->
    <aside class="pm-sidebar">
        <a class="pm-brand" href="<%=AppUrl("home.asp")%>">
            <span class="pm-brand-mark">
                <svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M9 18V5l12-2v13" /><circle cx="6" cy="18" r="3" /><circle cx="18" cy="16" r="3" />
                </svg>
            </span>
            <span class="pm-brand-text">
                <strong><%=H(SITE_NAME)%></strong>
                <span><%=H(SITE_TAGLINE)%></span>
            </span>
        </a>

        <nav class="pm-nav">
            <span class="pm-nav-label">Descubrir</span>

            <a class="<%=NavActivo("home")%>" href="<%=AppUrl("home.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M3 10.5 12 3l9 7.5" /><path d="M5 9.5V21h14V9.5" />
                </svg><span>Inicio</span>
            </a>

            <a class="<%=NavActivo("search")%>" href="<%=AppUrl("search.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <circle cx="11" cy="11" r="7" /><path d="m20 20-3.2-3.2" />
                </svg><span>Explorar</span>
            </a>

            <a class="<%=NavActivo("songs")%>" href="<%=AppUrl("songs.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M9 18V5l12-2v13" /><circle cx="6" cy="18" r="3" /><circle cx="18" cy="16" r="3" />
                </svg><span>Canciones</span>
            </a>

            <a class="<%=NavActivo("albums")%>" href="<%=AppUrl("albums.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9">
                    <circle cx="12" cy="12" r="9" /><circle cx="12" cy="12" r="2.6" />
                </svg><span>Albumes</span>
            </a>

            <a class="<%=NavActivo("artists")%>" href="<%=AppUrl("artists.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <circle cx="12" cy="8" r="4" /><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8" />
                </svg><span>Artistas</span>
            </a>

            <span class="pm-nav-label">Mi biblioteca</span>

            <a class="<%=NavActivo("playlists")%>" href="<%=AppUrl("playlists.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <path d="M4 6h16M4 12h16M4 18h10" />
                </svg><span>Playlists</span>
            </a>

            <a class="<%=NavActivo("profile")%>" href="<%=AppUrl("profile.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <circle cx="12" cy="8" r="4" /><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8" />
                </svg><span>Mi perfil</span>
            </a>

            <% If IsAdmin() Then %>
            <span class="pm-nav-label">Administracion</span>

            <a class="pm-nav-admin <%=NavActivo("admin-dashboard")%>" href="<%=AppUrl("admin/dashboard.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <path d="M4 13h6V4H4zM14 20h6v-9h-6zM4 20h6v-4H4zM14 8h6V4h-6z" />
                </svg><span>Panel</span>
            </a>

            <a class="pm-nav-admin <%=NavActivo("admin-users")%>" href="<%=AppUrl("admin/users.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <circle cx="9" cy="8" r="3.4" /><path d="M2.5 20c0-3.6 2.9-6.5 6.5-6.5s6.5 2.9 6.5 6.5" />
                    <path d="M17 11.5h4M19 9.5v4" />
                </svg><span>Usuarios</span>
            </a>

            <a class="pm-nav-admin <%=NavActivo("admin-songs")%>" href="<%=AppUrl("admin/songs.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M9 18V5l12-2v13" /><circle cx="6" cy="18" r="3" /><circle cx="18" cy="16" r="3" />
                </svg><span>Canciones</span>
            </a>

            <a class="pm-nav-admin <%=NavActivo("admin-albums")%>" href="<%=AppUrl("admin/albums.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9">
                    <circle cx="12" cy="12" r="9" /><circle cx="12" cy="12" r="2.6" />
                </svg><span>Albumes</span>
            </a>

            <a class="pm-nav-admin <%=NavActivo("admin-artists")%>" href="<%=AppUrl("admin/artists.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <circle cx="12" cy="8" r="4" /><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8" />
                </svg><span>Artistas</span>
            </a>

            <a class="pm-nav-admin <%=NavActivo("admin-playlists")%>" href="<%=AppUrl("admin/playlists.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <path d="M4 6h16M4 12h16M4 18h10" />
                </svg><span>Playlists</span>
            </a>

            <a class="pm-nav-admin <%=NavActivo("admin-catalog")%>" href="<%=AppUrl("admin/catalog.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round">
                    <path d="M4 5h16v4H4zM4 15h16v4H4z" /><path d="M8 9v6M16 9v6" />
                </svg><span>Catalogo</span>
            </a>
            <% End If %>
        </nav>

        <div class="pm-sidebar-foot">
            <strong style="color:var(--text-dim)">PoliMusic v1.0</strong><br />
            Repositorio musical &middot; BDD_PoliMusic
        </div>
    </aside>

    <!-- =============================================================== MAIN -->
    <main class="pm-main">
        <div class="pm-topbar">
            <form class="pm-search" method="get" action="<%=AppUrl("search.asp")%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <circle cx="11" cy="11" r="7" /><path d="m20 20-3.2-3.2" />
                </svg>
                <input type="search" name="q" placeholder="Buscar canciones, albumes, artistas..."
                       value="<%=H(gNavQ)%>" autocomplete="off" />
            </form>

            <div class="pm-topbar-right">
                <div class="pm-user">
                    <img class="pm-avatar" alt="" src="<%=MediaUrl(CurrentUserPhoto(), "user", CurrentUsername())%>" />
                    <span class="pm-user-meta">
                        <strong><%=H(CurrentUsername())%></strong>
                        <span><%=UCase(H(TipoUsuarioNombre(CurrentUserType())))%></span>
                    </span>
                </div>

                <a class="pm-btn pm-btn-sm" href="<%=AppUrl("logout.asp")%>" title="Cerrar sesion">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M15 17l5-5-5-5" /><path d="M20 12H9" /><path d="M12 20H5V4h7" />
                    </svg>
                    Salir
                </a>
            </div>
        </div>

        <%
        Dim gFlashMsg, gFlashTipo
        gFlashMsg = TakeFlash()
        gFlashTipo = TakeFlashTipo()
        If gFlashMsg <> "" Then
        %>
        <div class="pm-alert <%=IIf(gFlashTipo = "err", "pm-alert-err", "pm-alert-ok")%>"><%=H(gFlashMsg)%></div>
        <%
        End If
        %>
