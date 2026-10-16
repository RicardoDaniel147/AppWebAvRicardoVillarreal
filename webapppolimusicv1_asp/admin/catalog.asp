<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Catalogo generico TBL_CATALOG
' =====================================================================
RequireAdmin

Dim rsCatalog, cCount

Set rsCatalog = Query("SELECT ID_CATALOG, COD_CATALOG_PARENT, COD_CATALOG, VALUE FROM TBL_CATALOG " & _
                      "ORDER BY COD_CATALOG_PARENT, COD_CATALOG")
cCount = 0

gPage = "admin-catalog"
gTitle = "Catalogo"
%>
<!--#include file="../includes/header.asp" -->

<h1 class="pm-h1">Catalogo</h1>
<p class="pm-sub">Tabla generica <code>TBL_CATALOG</code>: define los tipos de usuario y de playlist.</p>

<!--#include file="_tabs.asp" -->

<div class="pm-section" style="margin-top:0">
    <h2 style="font-size:15px;color:var(--text-dim)">Entradas del catalogo</h2>
    <a class="pm-btn pm-btn-primary" href="<%=AppUrl("admin/catalog_edit.asp")%>">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round">
            <path d="M12 5v14M5 12h14" />
        </svg>
        Nueva entrada
    </a>
</div>

<% If rsCatalog.EOF Then %>
<div class="pm-empty"><strong>Catalogo vacio</strong>Agrega entradas para clasificar tipos de usuario y playlist.</div>
<% Else %>
<div class="pm-table-wrap">
    <table class="pm-table">
        <thead>
            <tr><th>ID</th><th>Padre</th><th>Codigo</th><th>Valor</th><th>Tipo</th><th style="text-align:right">Acciones</th></tr>
        </thead>
        <tbody>
            <% Dim cRoot
               Do While Not rsCatalog.EOF
                   cCount = cCount + 1
                   cRoot = (Trim("" & rsCatalog("COD_CATALOG_PARENT")) = "") %>
            <tr>
                <td><%=rsCatalog("ID_CATALOG")%></td>
                <td><%=IIf(cRoot, "-", H("" & rsCatalog("COD_CATALOG_PARENT")))%></td>
                <td><%=H("" & rsCatalog("COD_CATALOG"))%></td>
                <td><strong><%=H("" & rsCatalog("VALUE"))%></strong></td>
                <td>
                    <% If cRoot Then %>
                    <span class="pm-badge pm-badge-admin">Categoria raiz</span>
                    <% Else %>
                    <span class="pm-badge pm-badge-user">Detalle</span>
                    <% End If %>
                </td>
                <td>
                    <span class="actions">
                        <a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/catalog_edit.asp")%>?id=<%=rsCatalog("ID_CATALOG")%>">Editar</a>
                        <form action="<%=AppUrl("admin/delete.asp")%>" method="post" style="margin:0"
                              onsubmit="return confirm('Eliminar la entrada <%=JS(rsCatalog("VALUE"))%>?');">
                            <%=CsrfFieldInline()%>
                            <input type="hidden" name="entity" value="catalog" />
                            <input type="hidden" name="id" value="<%=rsCatalog("ID_CATALOG")%>" />
                            <button type="submit" class="pm-btn pm-btn-sm pm-btn-danger">Eliminar</button>
                        </form>
                    </span>
                </td>
            </tr>
            <% rsCatalog.MoveNext
               Loop %>
        </tbody>
    </table>
</div>
<% End If %>

<% rsCatalog.Close : Set rsCatalog = Nothing %>

<!--#include file="../includes/footer.asp" -->
