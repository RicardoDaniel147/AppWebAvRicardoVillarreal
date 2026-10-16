<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Detalle de una playlist
' =====================================================================
RequireLogin

Dim plId, rsPl, rsSongs, rsDisp, plCanEdit, plOwner, plType

plId = ToInt(Qry("id"), 0)
Set rsPl = Query(SqlPlaylists("pl.ID_PLAYLIST = " & plId, "", 0))

If rsPl.EOF Then
    rsPl.Close
    Set rsPl = Nothing
    Response.Redirect AppUrl("error.asp?code=404")
    Response.End
End If

plOwner   = ToInt(rsPl("ID_USER"), 0)
plType    = ToInt(rsPl("TYPE"), PLAYLIST_PUBLIC)
plCanEdit = PuedeEditarPlaylist(plId, plOwner, plType)

' Las playlists privadas solo las ve su dueno o un administrador.
If Not PuedeVerPlaylist(plOwner, plType) Then
    rsPl.Close
    Set rsPl = Nothing
    Response.Redirect AppUrl("error.asp?code=403")
    Response.End
End If

Set rsSongs = Query(SqlCanciones("EXISTS (SELECT 1 FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = s.ID_SONG AND ps.ID_PLAYLIST = " & plId & ")", "s.SONG_NAME", 0))

gPlaylistId = plId
gPage = "playlists"
gTitle = "" & rsPl("TITLE")
%>
<!--#include file="includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("playlists.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver a playlists
</a>

<section class="pm-detail pm-fade-in">
    <div class="pm-detail-art">
        <img src="<%=H(MediaUrl("" & rsPl("PHOTO"), "playlist", "" & rsPl("TITLE")))%>" alt="<%=H(rsPl("TITLE"))%>" />
    </div>

    <div>
        <span class="pm-eyebrow">Playlist</span>
        <h1><%=H(rsPl("TITLE"))%></h1>
        <div class="pm-detail-meta">
            <span class="<%=BadgePlaylist(plType)%>"><%=TipoPlaylistNombre(plType)%></span>
            <span>Creada por <%=H("" & rsPl("PROPIETARIO"))%></span>
            <span class="dot"></span>
            <span><%=FechaCorta(rsPl("CREATION_DATE"))%></span>
            <span class="dot"></span>
            <span><%=rsPl("CANCIONES")%> canciones</span>
        </div>

        <div class="pm-btn-row" style="margin-top:22px">
            <button type="button" class="pm-btn pm-btn-primary" onclick="PoliMusic.playFirst()">
                <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z" /></svg>
                Reproducir playlist
            </button>

            <% If plCanEdit Then %>
            <a class="pm-btn" href="<%=AppUrl("playlist_edit.asp")%>?id=<%=plId%>">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M12 20h9" /><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z" />
                </svg>
                Editar
            </a>

            <form action="<%=AppUrl("playlist_item.asp")%>" method="post" style="margin:0"
                  onsubmit="return confirm('Eliminar esta playlist? Esta accion no se puede deshacer.');">
                <%=CsrfFieldInline()%>
                <input type="hidden" name="action" value="delete" />
                <input type="hidden" name="id" value="<%=plId%>" />
                <button type="submit" class="pm-btn pm-btn-danger">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M3 6h18" /><path d="M8 6V4h8v2" /><path d="M19 6l-1 14H6L5 6" />
                    </svg>
                    Eliminar
                </button>
            </form>
            <% End If %>
        </div>
    </div>
</section>

<% rsPl.Close : Set rsPl = Nothing %>

<div class="pm-section">
    <h2>Canciones</h2>
</div>

<!--#include file="includes/songlist.asp" -->

<% rsSongs.Close : Set rsSongs = Nothing %>

<% If plCanEdit Then %>
<div class="pm-section">
    <h2>Agregar canciones</h2>
</div>

<div class="pm-panel">
    <% Set rsDisp = Query(SqlCanciones("NOT EXISTS (SELECT 1 FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = s.ID_SONG AND ps.ID_PLAYLIST = " & plId & ")", "s.SONG_NAME", 0)) %>

    <% If rsDisp.EOF Then %>
    <p class="pm-hint" style="margin:0">Todas las canciones del catalogo ya estan en esta playlist.</p>
    <% Else %>
    <form action="<%=AppUrl("playlist_item.asp")%>" method="post" class="pm-form">
        <%=CsrfFieldInline()%>
        <input type="hidden" name="action" value="add" />
        <input type="hidden" name="id" value="<%=plId%>" />

        <div class="pm-field">
            <label for="song">Cancion</label>
            <select id="song" name="song" class="pm-select">
                <% Do While Not rsDisp.EOF %>
                <option value="<%=rsDisp("ID_SONG")%>"><%=H(rsDisp("SONG_NAME"))%> &mdash; <%=H("" & rsDisp("ARTISTAS"))%></option>
                <% rsDisp.MoveNext
                   Loop %>
            </select>
        </div>

        <div class="pm-btn-row">
            <button type="submit" class="pm-btn pm-btn-primary">Agregar a la playlist</button>
        </div>
    </form>
    <% End If %>

    <% rsDisp.Close : Set rsDisp = Nothing %>
</div>
<% End If %>

<!--#include file="includes/footer.asp" -->
