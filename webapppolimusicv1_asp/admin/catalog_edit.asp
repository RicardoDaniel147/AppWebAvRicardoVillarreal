<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="../includes/common.asp" -->
<%
' =====================================================================
'  PoliMusic - Crear / editar entrada del catalogo (solo administradores)
' =====================================================================
RequireAdmin

Dim ceId, ceNew, ceParent, ceCod, ceValue, ceErr, rsC

ceId     = ToInt(Qry("id"), 0)
ceNew    = (ceId = 0)
ceErr    = ""
ceParent = ""
ceCod    = ""
ceValue  = ""

If Not ceNew Then
    Set rsC = Query("SELECT ID_CATALOG, COD_CATALOG_PARENT, COD_CATALOG, VALUE FROM TBL_CATALOG WHERE ID_CATALOG = " & ceId)
    If rsC.EOF Then
        rsC.Close : Set rsC = Nothing
        Response.Redirect AppUrl("error.asp?code=404")
        Response.End
    End If
    ceParent = "" & rsC("COD_CATALOG_PARENT")
    ceCod    = "" & rsC("COD_CATALOG")
    ceValue  = "" & rsC("VALUE")
    rsC.Close : Set rsC = Nothing
End If

If UCase("" & Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    ParseMultipart
    CsrfCheck()

    ceParent = PostVal("codparent")
    ceCod    = PostVal("codcatalog")
    ceValue  = PostVal("value")

    If Len(Trim(ceValue)) = 0 Then
        ceErr = "El valor es obligatorio."
    ElseIf Len(ceParent) > 8 Then
        ceErr = "El codigo padre no puede superar los 8 caracteres."
    ElseIf Len(ceCod) > 8 Then
        ceErr = "El codigo no puede superar los 8 caracteres."
    End If

    If ceErr = "" Then
        If ceNew Then
            ExecuteInsert "INSERT INTO TBL_CATALOG(COD_CATALOG_PARENT, COD_CATALOG, VALUE) VALUES(" & _
                SqlText(Trim(ceParent)) & ", " & SqlText(Trim(ceCod)) & ", " & SqlText(Trim(ceValue)) & ")"
            Flash "Entrada de catalogo creada correctamente.", "ok"
        Else
            Execute "UPDATE TBL_CATALOG SET COD_CATALOG_PARENT = " & SqlText(Trim(ceParent)) & _
                    ", COD_CATALOG = " & SqlText(Trim(ceCod)) & _
                    ", VALUE = " & SqlText(Trim(ceValue)) & _
                    " WHERE ID_CATALOG = " & ceId
            Flash "Entrada de catalogo actualizada.", "ok"
        End If

        Response.Redirect AppUrl("admin/catalog.asp")
        Response.End
    End If
End If

gPage = "admin-catalog"
gTitle = IIf(ceNew, "Nueva entrada de catalogo", "Editar entrada de catalogo")
%>
<!--#include file="../includes/header.asp" -->

<a class="pm-btn pm-btn-sm" href="<%=AppUrl("admin/catalog.asp")%>" style="margin-bottom:18px">
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M15 18l-6-6 6-6" />
    </svg>
    Volver al catalogo
</a>

<h1 class="pm-h1"><%=H(gTitle)%></h1>
<p class="pm-sub">Codigo padre vacio = categoria raiz. Con padre = valor de detalle de esa categoria.</p>

<% If ceErr <> "" Then %>
<div class="pm-alert pm-alert-err"><%=H(ceErr)%></div>
<% End If %>

<form method="post" action="<%=AppUrl("admin/catalog_edit.asp")%>?id=<%=ceId%>" class="pm-form">
    <%=CsrfFieldInline()%>

    <div class="pm-panel">
        <h3>Entrada</h3>

        <div class="pm-form-grid">
            <div class="pm-field">
                <label for="codparent">Codigo padre</label>
                <input type="text" id="codparent" name="codparent" class="pm-input" maxlength="8"
                       value="<%=H(ceParent)%>" placeholder="(vacio para categoria raiz)" />
                <span class="pm-hint">Ejemplo: 1 para tipos de usuario, 2 para tipos de playlist.</span>
            </div>

            <div class="pm-field">
                <label for="codcatalog">Codigo</label>
                <input type="text" id="codcatalog" name="codcatalog" class="pm-input" maxlength="8"
                       value="<%=H(ceCod)%>" placeholder="0, 1, 2..." />
            </div>

            <div class="pm-field pm-span-2">
                <label for="value">Valor</label>
                <input type="text" id="value" name="value" class="pm-input" maxlength="255"
                       value="<%=H(ceValue)%>" placeholder="USER TYPE, ADMIN, Public, Private..." />
            </div>
        </div>
    </div>

    <div class="pm-btn-row">
        <button type="submit" class="pm-btn pm-btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <path d="M20 6 9 17l-5-5" />
            </svg>
            <%=IIf(ceNew, "Crear entrada", "Guardar cambios")%>
        </button>
        <a class="pm-btn" href="<%=AppUrl("admin/catalog.asp")%>">Cancelar</a>
    </div>
</form>

<!--#include file="../includes/footer.asp" -->
