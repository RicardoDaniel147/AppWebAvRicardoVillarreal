<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Playlists del usuario y playlists publicas
' =====================================================================
RequireLogin

Dim rsMine, rsPublic, plCount

Set rsMine   = Query(SqlPlaylists("pl.ID_USER = " & CurrentUserId(), "pl.CREATION_DATE DESC", 0))
Set rsPublic = Query(SqlPlaylists("pl.TYPE = 0", "pl.CREATION_DATE DESC", 12))
plCount = 0

gPage = "playlists"
gTitle = "Playlists"
%>
<!--#include file="includes/header.asp" -->

<div class="pm-section" style="margin-top:0">
    <div>
        <h1 class="pm-h1" style="margin-bottom:2px">Playlists</h1>
        <p class="pm-sub" style="margin:0">Organiza tu musica en colecciones.</p>
    </div>
    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("playlist_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nueva playlist
    </a>
</div>

<div class="pm-section">
    <h2>Mis playlists</h2>
</div>

<% If rsMine.EOF Then %>
<div class="pm-empty">
    <strong>Todavia no tienes playlists</strong>
    Crea tu primera playlist y empieza a agrupar tus canciones favoritas.
</div>
<% Else %>
<div class="pm-grid">
    <% Do While Not rsMine.EOF
           plCount = plCount + 1 %>
    <a class="pm-card" href="<%=AppUrl("playlist.asp")%>?id=<%=rsMine("ID_PLAYLIST")%>">
        <div class="pm-card-art">
            <img src="<%=H(MediaUrl("" & rsMine("PHOTO"), "playlist", "" & rsMine("TITLE")))%>" alt="" loading="lazy" />
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsMine("TITLE"))%></div>
            <div class="pm-card-meta">
                <%=rsMine("CANCIONES")%> canciones &middot;
                <span class="<%=BadgePlaylist(rsMine("TYPE"))%>"><%=TipoPlaylistNombre(rsMine("TYPE"))%></span>
            </div>
        </div>
    </a>
    <% rsMine.MoveNext
       Loop %>
</div>
<% End If %>
<% rsMine.Close : Set rsMine = Nothing %>

<div class="pm-section">
    <h2>Playlists publicas de la comunidad</h2>
</div>

<% If rsPublic.EOF Then %>
<div class="pm-empty"><strong>Sin playlists publicas</strong>Aun no hay playlists compartidas.</div>
<% Else %>
<div class="pm-grid">
    <% Do While Not rsPublic.EOF %>
    <a class="pm-card" href="<%=AppUrl("playlist.asp")%>?id=<%=rsPublic("ID_PLAYLIST")%>">
        <div class="pm-card-art">
            <img src="<%=H(MediaUrl("" & rsPublic("PHOTO"), "playlist", "" & rsPublic("TITLE")))%>" alt="" loading="lazy" />
        </div>
        <div>
            <div class="pm-card-title"><%=H(rsPublic("TITLE"))%></div>
            <div class="pm-card-meta"><%=rsPublic("CANCIONES")%> canciones &middot; <%=H("" & rsPublic("PROPIETARIO"))%></div>
        </div>
    </a>
    <% rsPublic.MoveNext
       Loop %>
</div>
<% End If %>
<% rsPublic.Close : Set rsPublic = Nothing %>

<!--#include file="includes/footer.asp" -->
