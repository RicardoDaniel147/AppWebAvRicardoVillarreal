<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Administracion de canciones
' =====================================================================
RequireAdmin

Dim sgSearch, sgGenre, sgArtist, sgCond, sgGenres, rsSongs, rsArtistas, sgi

sgSearch = Qry("q")
sgGenre  = Qry("genre")
sgArtist = ToInt(Qry("artistId"), 0)

sgCond = ""
If sgSearch <> "" Then sgCond = sgCond & IIf(sgCond = "", "", " AND ") & "s.SONG_NAME LIKE '%' + " & SqlText(sgSearch) & " + '%'"
If sgGenre <> ""  Then sgCond = sgCond & IIf(sgCond = "", "", " AND ") & "s.GENRE = " & SqlText(sgGenre)
If sgArtist > 0   Then sgCond = sgCond & IIf(sgCond = "", "", " AND ") & "EXISTS (SELECT 1 FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG AND p.ID_ARTIST = " & sgArtist & ")"

Set rsSongs    = Query(SqlCanciones(sgCond, "s.SONG_NAME", 0))
sgGenres       = GenerosCanciones()
Set rsArtistas = Query(SqlArtistas("", "a.ARTIST_NAME", 0))

gPage = "admin-songs"
gTitle = "Administrar canciones"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Canciones</h1>
<p class="pm-sub">Gestiona el audio, el genero y los artistas de cada pista.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-section" style="margin-top:0">
    <form method="get" action="<%=AppUrl("admin/songs.asp")%>" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <input type="search" name="q" class="pm-input" style="width:220px" value="<%=H(sgSearch)%>" placeholder="Buscar cancion..." />
        <select name="genre" class="pm-select" style="width:170px">
            <option value="">Todos los generos</option>
            <% If UBound(sgGenres) >= 0 Then
                   For sgi = 0 To UBound(sgGenres) %>
            <option value="<%=H(sgGenres(sgi))%>"<%=SelIf(sgGenres(sgi), sgGenre)%>><%=H(sgGenres(sgi))%></option>
            <%     Next
               End If %>
        </select>
        <select name="artistId" class="pm-select" style="width:190px">
            <option value="0">Todos los artistas</option>
            <% Do While Not rsArtistas.EOF %>
            <option value="<%=rsArtistas("ID_ARTIST")%>"<%=IIf(sgArtist = ToInt(rsArtistas("ID_ARTIST"), 0), " selected=""selected""", "")%>><%=H(NombreArtista(rsArtistas("ARTIST_NAME"), rsArtistas("ARTIST_LASTNAME")))%></option>
            <% rsArtistas.MoveNext
               Loop %>
        </select>
        <button type="submit" class="pm-btn pm-btn-sm">Filtrar</button>
        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/songs.asp")%>">Limpiar</a>
    </form>

    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/song_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nueva cancion
    </a>
</div>

<% If rsSongs.EOF Then %>
<div class="pm-empty"><strong>Sin canciones</strong>No hay pistas que coincidan con el filtro.</div>
<% Else %>
<div class="pm-table-wrap">
    <table class="pm-table">
        <thead>
            <tr>
                <th>#</th><th>Cancion</th><th>Artistas</th><th>Genero</th>
                <th>Reproducciones</th><th>Archivo</th><th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Dim sgOk
               Do While Not rsSongs.EOF
                   sgOk = AudioExists("" & rsSongs("SONG_PATH")) %>
            <tr>
                <td><%=rsSongs("ID_SONG")%></td>
                <td><strong><%=H(rsSongs("SONG_NAME"))%></strong></td>
                <td><%=H("" & rsSongs("ARTISTAS"))%></td>
                <td><span class="pm-badge pm-badge-user"><%=H("" & rsSongs("GENRE"))%></span></td>
                <td><%=Num(rsSongs("PLAYS"))%></td>
                <td>
                    <% If sgOk Then %>
                    <span class="pm-badge pm-badge-public">Disponible</span>
                    <% Else %>
                    <span class="pm-badge pm-badge-warn">Falta</span>
                    <% End If %>
                </td>
                <td>
                    <span class="actions">
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/song_edit.asp")%>?id=<%=rsSongs("ID_SONG")%>">Editar</a>
                        <form action="<%=AppUrl("admin/delete.asp")%>" method="post" style="margin:0"
                              onsubmit="return confirm('Eliminar la cancion <%=JS(rsSongs("SONG_NAME"))%>? Se quitara de albumes y playlists.');">
                            <%=CsrfFieldInline()%>
                            <input type="hidden" name="entity" value="song" />
                            <input type="hidden" name="id" value="<%=rsSongs("ID_SONG")%>" />
                            <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger">Eliminar</button>
                        </form>
                    </span>
                </td>
            </tr>
            <% rsSongs.MoveNext
               Loop %>
        </tbody>
    </table>
</div>
<% End If %>

<% rsSongs.Close : Set rsSongs = Nothing %>
<% rsArtistas.Close : Set rsArtistas = Nothing %>

<!--#include file="../includes/footer.asp" -->
