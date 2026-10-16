<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Administracion de usuarios
' =====================================================================
RequireAdmin

Dim uSearch, uCond, rsUsers, uCount

uSearch = Qry("q")
uCond = ""
If uSearch <> "" Then
    uCond = "USERNAME LIKE '%' + " & SqlText(uSearch) & " + '%' OR EMAIL LIKE '%' + " & SqlText(uSearch) & " + '%'"
End If

Set rsUsers = Query("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER " & _
                    IIf(uCond = "", "", "WHERE " & uCond) & " ORDER BY USER_TYPE DESC, USERNAME")
uCount = 0

gPage = "admin-users"
gTitle = "Administrar usuarios"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Usuarios</h1>
<p class="pm-sub">Solo el administrador puede crear, editar o eliminar cuentas.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-section" style="margin-top:0">
    <form method="get" action="<%=AppUrl("admin/users.asp")%>" style="display:flex;gap:10px;align-items:center">
        <input type="search" name="q" class="pm-input" style="width:280px" value="<%=H(uSearch)%>"
               placeholder="Buscar por usuario o correo..." />
        <button type="submit" class="pm-btn pm-btn-sm">Buscar</button>
        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/users.asp")%>">Limpiar</a>
    </form>

    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/user_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nuevo usuario
    </a>
</div>

<% If rsUsers.EOF Then %>
<div class="pm-empty"><strong>Sin usuarios</strong>No hay cuentas que coincidan con la busqueda.</div>
<% Else %>
<div class="pm-table-wrap">
    <table class="pm-table">
        <thead>
            <tr>
                <th>Usuario</th>
                <th>Correo</th>
                <th>Tipo</th>
                <th>Nacimiento</th>
                <th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Do While Not rsUsers.EOF
                   uCount = uCount + 1 %>
            <tr>
                <td>
                    <span class="cell-main">
                        <img class="round" src="<%=H(MediaUrl("" & rsUsers("USER_PHOTO"), "user", "" & rsUsers("USERNAME")))%>" alt="" />
                        <span>
                            <strong><%=H(rsUsers("USERNAME"))%></strong>
                            <span class="pm-hint" style="display:block">ID <%=rsUsers("ID_USER")%></span>
                        </span>
                    </span>
                </td>
                <td><%=H(rsUsers("EMAIL"))%></td>
                <td><span class="<%=BadgeUsuario(ToInt(rsUsers("USER_TYPE"), 0) = USER_TYPE_ADMIN)%>"><%=TipoUsuarioNombre(rsUsers("USER_TYPE"))%></span></td>
                <td><%=FechaCorta(rsUsers("BIRTHDAY"))%></td>
                <td>
                    <span class="actions">
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/user_edit.asp")%>?id=<%=rsUsers("ID_USER")%>">Editar</a>
                        <form action="<%=AppUrl("admin/delete.asp")%>" method="post" style="margin:0"
                              onsubmit="return confirm('Eliminar el usuario <%=JS(rsUsers("USERNAME"))%>? Tambien se eliminaran sus playlists.');">
                            <%=CsrfFieldInline()%>
                            <input type="hidden" name="entity" value="user" />
                            <input type="hidden" name="id" value="<%=rsUsers("ID_USER")%>" />
                            <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger">Eliminar</button>
                        </form>
                    </span>
                </td>
            </tr>
            <% rsUsers.MoveNext
               Loop %>
        </tbody>
    </table>
</div>
<% End If %>

<% rsUsers.Close : Set rsUsers = Nothing %>

<!--#include file="../includes/footer.asp" -->
