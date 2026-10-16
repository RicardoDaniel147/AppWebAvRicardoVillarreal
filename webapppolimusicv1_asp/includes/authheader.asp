<%
' =====================================================================
'  PoliMusic - Maqueta de las pantallas de acceso (login / registro)
'  La pagina debe asignar gTitle antes de incluir este archivo.
' =====================================================================
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title><%=H(gTitle)%> &middot; <%=H(SITE_NAME)%></title>
    <link rel="stylesheet" href="<%=Asset("style.css")%>" />
</head>
<body class="pm-scroll">
<div class="pm-bg"></div>

<div class="pm-auth">
    <section class="pm-auth-art">
        <a class="pm-brand" href="<%=AppUrl("default.asp")%>" style="padding:0">
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

        <h1>Tu musica.<br /><em>Tu control.</em></h1>
        <p>
            Plataforma de streaming musical con catalogo administrable, playlists personalizadas
            y un reproductor continuo. Construida en ASP clasico con SQL Server.
        </p>

        <div class="pm-auth-features">
            <div class="pm-auth-feature">
                <i><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M9 18V5l12-2v13" /><circle cx="6" cy="18" r="3" /><circle cx="18" cy="16" r="3" /></svg></i>
                <span>Reproductor continuo con cola, aleatorio y repeticion</span>
            </div>
            <div class="pm-auth-feature">
                <i><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <circle cx="11" cy="11" r="7" /><path d="m20 20-3.2-3.2" /></svg></i>
                <span>Filtra el catalogo por genero, artista y album</span>
            </div>
            <div class="pm-auth-feature">
                <i><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <path d="M4 6h16M4 12h16M4 18h10" /></svg></i>
                <span>Crea y administra tus propias playlists</span>
            </div>
            <div class="pm-auth-feature">
                <i><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                    <path d="M12 3l8 4v6c0 4.4-3.4 8-8 8s-8-3.6-8-8V7z" /></svg></i>
                <span>Roles diferenciados: administrador y usuario</span>
            </div>
        </div>
    </section>

    <section class="pm-auth-panel">
        <div class="pm-auth-card pm-fade-in">
