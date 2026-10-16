<%
' =====================================================================
'  PoliMusic - Listado de canciones reutilizable
'  Requiere que la pagina haya creado el Recordset "rsSongs" con las
'  columnas: ID_SONG, SONG_NAME, SONG_PATH, PLAYS, GENRE, ARTISTAS, ARTIST_IMAGE
'  Opcionalmente gPlaylistId para mostrar el boton de quitar.
' =====================================================================
Dim slIdx, slSrc, slExists, slArt, slCover, slGenre, slPlays
%>
<% If rsSongs.EOF Then %>
<div class="pm-empty">
    <strong>No hay canciones que mostrar</strong>
    Ajusta los filtros o agrega contenido desde el panel de administracion.
</div>
<% Else %>
<div data-playlist-scope>
    <div class="pm-list-head">
        <span>#</span>
        <span>Titulo</span>
        <span class="h-genre">Genero</span>
        <span class="h-plays">Reproducciones</span>
        <span></span>
    </div>

    <div class="pm-list">
    <%
    slIdx = 0
    Do While Not rsSongs.EOF
        slIdx = slIdx + 1
        slSrc = AudioUrl("" & rsSongs("SONG_PATH"))
        slExists = (Len(slSrc) > 0)
        slArt = "" & rsSongs("ARTISTAS")
        slCover = MediaUrl("" & rsSongs("ARTIST_IMAGE"), "artist", slArt)
        slGenre = "" & rsSongs("GENRE")
        slPlays = ToInt(rsSongs("PLAYS"), 0)
        If slGenre = "" Then slGenre = "Sin genero"
    %>
        <div class="pm-row"
             data-song-id="<%=rsSongs("ID_SONG")%>"
             data-song-name="<%=H(rsSongs("SONG_NAME"))%>"
             data-song-artists="<%=H(slArt)%>"
             data-song-src="<%=H(slSrc)%>"
             data-song-cover="<%=H(slCover)%>"
             data-song-exists="<%=IIf(slExists, "1", "0")%>"
             title="<%=IIf(slExists, "Reproducir", "Archivo de audio no disponible")%>">

            <span class="pm-row-idx"><%=slIdx%></span>

            <span class="pm-row-main">
                <img class="pm-row-thumb" src="<%=H(slCover)%>" alt="" loading="lazy" />
                <span class="pm-row-text">
                    <span class="pm-row-name"><%=H(rsSongs("SONG_NAME"))%></span>
                    <span class="pm-row-sub"><%=H(slArt)%></span>
                </span>
            </span>

            <span class="pm-row-genre"><%=H(slGenre)%></span>
            <span class="pm-row-plays"><%=Num(slPlays)%></span>

            <span class="pm-row-actions">
                <% If gPlaylistId > 0 Then %>
                <form action="<%=AppUrl("playlist_item.asp")%>" method="post" class="pm-no-play" style="margin:0">
                    <%=CsrfFieldInline()%>
                    <input type="hidden" name="action" value="remove" />
                    <input type="hidden" name="id" value="<%=gPlaylistId%>" />
                    <input type="hidden" name="song" value="<%=rsSongs("ID_SONG")%>" />
                    <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger" title="Quitar de la playlist"
                            onclick="return confirm('Quitar esta cancion de la playlist?');">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
                            <path d="M5 12h14" />
                        </svg>
                    </button>
                </form>
                <% ElseIf Not slExists Then %>
                <span class="pm-badge pm-badge-warn">Sin archivo</span>
                <% End If %>
            </span>
        </div>
    <%
        rsSongs.MoveNext
    Loop
    %>
    </div>
</div>
<% End If %>
