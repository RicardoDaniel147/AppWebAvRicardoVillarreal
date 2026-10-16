<%
' =====================================================================
'  PoliMusic - Sesion y control de permisos
' =====================================================================

Function IsLoggedIn()
    IsLoggedIn = (Len("" & Session("UserId")) > 0)
End Function

Function CurrentUserId()
    CurrentUserId = ToInt(Session("UserId"), 0)
End Function

Function CurrentUsername()
    CurrentUsername = "" & Session("Username")
End Function

Function CurrentUserPhoto()
    CurrentUserPhoto = "" & Session("UserPhoto")
End Function

Function CurrentUserEmail()
    CurrentUserEmail = "" & Session("Email")
End Function

Function CurrentUserType()
    CurrentUserType = ToInt(Session("UserType"), USER_TYPE_NORMAL)
End Function

Function IsAdmin()
    IsAdmin = (CurrentUserType() = USER_TYPE_ADMIN)
End Function

' Guarda la identidad del usuario en la sesion a partir de un Recordset.
Sub LoginSession(rs)
    Session("UserId")    = CLng(rs("ID_USER").Value)
    Session("Username")  = "" & rs("USERNAME")
    Session("UserType")  = CLng(rs("USER_TYPE").Value)
    Session("Email")     = "" & rs("EMAIL")
    Session("UserPhoto") = "" & rs("USER_PHOTO")
End Sub

' Cierra la sesion por completo.
Sub LogoutSession()
    Session.Contents.RemoveAll()
    Session.Abandon
End Sub

' Vuelve a leer el usuario desde la base de datos y actualiza la sesion.
Sub RefreshSession()
    Dim rs
    If Not IsLoggedIn() Then Exit Sub
    Set rs = Query("SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO " & _
                   "FROM TBL_USER WHERE ID_USER = " & CurrentUserId())
    If Not rs.EOF Then
        Session("Username")  = "" & rs("USERNAME")
        Session("UserType")  = CLng(rs("USER_TYPE").Value)
        Session("Email")     = "" & rs("EMAIL")
        Session("UserPhoto") = "" & rs("USER_PHOTO")
    End If
    rs.Close
    Set rs = Nothing
End Sub

' Exige sesion iniciada; si no, manda al login recordando el destino.
Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect AppPath() & "/login.asp?next=" & Server.URLEncode("" & Request.ServerVariables("URL"))
        Response.End
    End If
End Sub

' Exige rol de administrador.
Sub RequireAdmin()
    RequireLogin()
    If Not IsAdmin() Then
        Response.Redirect AppPath() & "/error.asp?code=403"
        Response.End
    End If
End Sub

' Comprueba que la playlist pertenece al usuario actual o que es administrador.
Function PuedeEditarPlaylist(idPlaylist, idPropietario, esPublica)
    PuedeEditarPlaylist = False
    If IsAdmin() Then
        PuedeEditarPlaylist = True
    ElseIf CurrentUserId() = ToInt(idPropietario, 0) Then
        PuedeEditarPlaylist = True
    ElseIf ToInt(esPublica, 0) = PLAYLIST_PUBLIC Then
        PuedeEditarPlaylist = False
    End If
End Function

' Comprueba si el usuario puede VER una playlist.
Function PuedeVerPlaylist(idPropietario, esPublica)
    PuedeVerPlaylist = True
    If ToInt(esPublica, 0) = PLAYLIST_PRIVATE Then
        If IsAdmin() Then
            PuedeVerPlaylist = True
        Else
            PuedeVerPlaylist = (CurrentUserId() = ToInt(idPropietario, 0))
        End If
    End If
End Function

' =====================================================================
'  Proteccion CSRF
'  Todas las paginas que modifican datos escriben CsrfField() dentro del
'  formulario y llaman a CsrfCheck() al principio del POST.
' =====================================================================

Function NuevoToken()
    Dim s, i
    Randomize Timer
    s = ""
    For i = 1 To 32
        s = s & Mid("0123456789abcdef", Int(Rnd() * 16) + 1, 1)
    Next
    NuevoToken = s & Len(Session.SessionID)
End Function

' Devuelve (y crea si hace falta) el token CSRF de la sesion.
Function CsrfToken()
    Dim t
    t = "" & Session("csrf")
    If Len(t) = 0 Then
        t = NuevoToken()
        Session("csrf") = t
    End If
    CsrfToken = t
End Function

' Escribe el campo oculto con el token.
Sub CsrfField()
    Response.Write "<input type=""hidden"" name=""_t"" value=""" & H(CsrfToken()) & """ />"
End Sub

' Devuelve el campo oculto como cadena de texto (util dentro de una expresion).
Function CsrfFieldInline()
    CsrfFieldInline = "<input type=""hidden"" name=""_t"" value=""" & H(CsrfToken()) & """ />"
End Function

' Valida el token en peticiones POST. Si no coincide, corta la ejecucion.
Sub CsrfCheck()
    Dim t
    If UCase("" & Request.ServerVariables("REQUEST_METHOD")) <> "POST" Then Exit Sub

    t = ""
    If IsMultipart() Then
        If IsObject(gFields) Then
            If gFields.Exists("_t") Then t = "" & gFields("_t")
        End If
    Else
        t = "" & Request.Form("_t")
    End If

    If Len(t) = 0 Or t <> ("" & Session("csrf")) Then
        Response.Status = "403 Forbidden"
        Response.Redirect AppPath() & "/error.asp?code=403"
        Response.End
    End If
End Sub

%>
