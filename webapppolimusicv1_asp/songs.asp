<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Listado de canciones con filtros
' =====================================================================
RequireLogin

Dim sSearch, sGenre, sArtist, sOrder, sCond, sOrderSql, sGenres, rsSongs, rsArtistas
Dim si, sq

sSearch = Qry("q")
sGenre  = Qry("genre")
sArtist = ToInt(Qry("artistId"), 0)
sOrder  = LCase(Qry("order"))
If sOrder = "" Then sOrder = "name"

Select Case sOrder
    Case "plays" : sOrderSql = "s.PLAYS DESC, s.SONG_NAME"
    Case "genre" : sOrderSql = "s.GENRE, s.SONG_NAME"
    Case Else    : sOrderSql = "s.SONG_NAME"
End Select

sCond = ""
If sSearch <> "" Then sCond = sCond & IIf(sCond = "", "", " AND ") & "s.SONG_NAME LIKE '%' + " & SqlText(sSearch) & " + '%'"
If sGenre <> ""  Then sCond = sCond & IIf(sCond = "", "", " AND ") & "s.GENRE = " & SqlText(sGenre)
If sArtist > 0   Then sCond = sCond & IIf(sCond = "", "", " AND ") & _
                              "EXISTS (SELECT 1 FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG AND p.ID_ARTIST = " & sArtist & ")"

Set rsSongs = Query(SqlCanciones(sCond, sOrderSql, 0))
sGenres = GenerosCanciones()
Set rsArtistas = Query(SqlArtistas("", "a.ARTIST_NAME", 0))

gPage = "songs"
gTitle = "Canciones"
%>
<!--#include file="includes/header.asp" -->

<h1 class="pm-h1">Canciones</h1>
<p class="pm-sub">Filtra el catalogo por texto, genero o artista.</p>

<div class="pm-panel" style="margin-bottom:22px">
    <form method="get" action="<%=AppUrl("songs.asp")%>" class="pm-form-grid">
        <div class="pm-field pm-span-2">
            <label for="q">Buscar</label>
            <input type="search" id="q" name="q" class="pm-input" value="<%=H(sSearch)%>"
                   placeholder="Nombre de la cancion..." autocomplete="off" />
        </div>

        <div class="pm-field">
            <label for="genre">Genero</label>
            <select id="genre" name="genre" class="pm-select">
                <option value="">Todos los generos</option>
                <% If UBound(sGenres) >= 0 Then
                       For si = 0 To UBound(sGenres) %>
                <option value="<%=H(sGenres(si))%>"<%=SelIf(sGenres(si), sGenre)%>><%=H(sGenres(si))%></option>
                <%     Next
                   End If %>
            </select>
        </div>

        <div class="pm-field">
            <label for="artistId">Artista</label>
            <select id="artistId" name="artistId" class="pm-select">
                <option value="0">Todos los artistas</option>
                <% Do While Not rsArtistas.EOF %>
                <option value="<%=rsArtistas("ID_ARTIST")%>"<%=IIf(sArtist = ToInt(rsArtistas("ID_ARTIST"), 0), " selected=""selected""", "")%>><%=H(NombreArtista(rsArtistas("ARTIST_NAME"), rsArtistas("ARTIST_LASTNAME")))%></option>
                <% rsArtistas.MoveNext
                   Loop %>
            </select>
        </div>

        <div class="pm-field">
            <label for="order">Ordenar por</label>
            <select id="order" name="order" class="pm-select">
                <option value="name"<%=SelIf("name", sOrder)%>>Nombre</option>
                <option value="plays"<%=SelIf("plays", sOrder)%>>Mas reproducidas</option>
                <option value="genre"<%=SelIf("genre", sOrder)%>>Genero</option>
            </select>
        </div>

        <div class="pm-btn-row" style="align-self:end">
            <button type="submit" class="pm-btn pm-btn-primary">Aplicar filtros</button>
            <a class="pm-btn" href="<%=AppUrl("songs.asp")%>">Limpiar</a>
        </div>
    </form>
</div>

<!--#include file="includes/songlist.asp" -->

<% rsSongs.Close : Set rsSongs = Nothing %>
<% rsArtistas.Close : Set rsArtistas = Nothing %>

<!--#include file="includes/footer.asp" -->
