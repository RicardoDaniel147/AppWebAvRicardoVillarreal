<%
' =====================================================================
'  PoliMusic - Acceso a datos (ADO)
'  Abre la conexion a BDD_PoliMusic probando varios proveedores OLE DB,
'  de modo que la aplicacion funcione en cualquier servidor Windows.
' =====================================================================

' Devuelve la lista de cadenas de conexion candidatas, en orden de preferencia.
Function ConnectionCandidates()
    Dim a(4)
    a(0) = "Provider=SQLOLEDB;Data Source=" & DB_SERVER & ";Initial Catalog=" & DB_NAME & _
           ";User ID=" & DB_USER & ";Password=" & DB_PASS & ";"
    a(1) = "Provider=SQLOLEDB;Data Source=" & DB_SERVER & ";Initial Catalog=" & DB_NAME & _
           ";User ID=" & DB_USER & ";Password=" & DB_PASS & ";Encrypt=True;TrustServerCertificate=True;"
    a(2) = "Provider=MSOLEDBSQL;Server=" & DB_SERVER & ";Database=" & DB_NAME & _
           ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";Encrypt=Optional;TrustServerCertificate=Yes;"
    a(3) = "Provider=MSDASQL;Driver={ODBC Driver 17 for SQL Server};Server=" & DB_SERVER & _
           ";Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & _
           ";Encrypt=yes;TrustServerCertificate=yes;"
    a(4) = "Provider=MSDASQL;Driver={ODBC Driver 18 for SQL Server};Server=" & DB_SERVER & _
           ";Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & _
           ";Encrypt=yes;TrustServerCertificate=yes;"
    ConnectionCandidates = a
End Function

' Comprueba si una cadena de conexion funciona realmente.
Function ConnectionWorks(cs)
    Dim cn
    ConnectionWorks = False
    On Error Resume Next
    Set cn = Server.CreateObject("ADODB.Connection")
    If Err.Number = 0 Then
        cn.CommandTimeout = 10
        cn.Open cs
        If Err.Number = 0 Then ConnectionWorks = True
        If Err.Number = 0 Then cn.Close
        Set cn = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Devuelve la cadena de conexion activa (se detecta una vez y se cachea).
Function ActiveConnString()
    Dim cached, list, i
    cached = "" & Application("POLIMUSIC_CS")
    If Len(cached) > 0 Then
        ActiveConnString = cached
        Exit Function
    End If

    list = ConnectionCandidates()
    For i = 0 To UBound(list)
        If ConnectionWorks(list(i)) Then
            Application.Lock
            Application("POLIMUSIC_CS") = list(i)
            Application.Unlock
            ActiveConnString = list(i)
            Exit Function
        End If
    Next

    ' Si ninguna funciono se devuelve la primera para que el error sea explicito.
    ActiveConnString = list(0)
End Function

' Invalida la cadena cacheada (util tras un error de conexion).
Sub ResetConnString()
    Application.Lock
    Application("POLIMUSIC_CS") = ""
    Application.Unlock
End Sub

' Abre una conexion nueva.
Function GetConn()
    Dim cn
    Set cn = Server.CreateObject("ADODB.Connection")
    cn.CommandTimeout = 30
    cn.Open ActiveConnString()
    Set GetConn = cn
End Function

' Cierra una conexion de forma segura.
Sub CloseConn(cn)
    On Error Resume Next
    If IsObject(cn) Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' Ejecuta un SELECT y devuelve un Recordset desconectado.
Function Query(sql)
    Dim cn, rs
    Set cn = GetConn()
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open sql, cn, 3, 1  ' adOpenStatic, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    CloseConn(cn)
    Set Query = rs
End Function

' Ejecuta un SELECT y devuelve el primer valor de la primera fila (o Empty).
Function Scalar(sql)
    Dim rs
    Set rs = Query(sql)
    If rs.EOF Then
        Scalar = Empty
    Else
        Scalar = rs.Fields(0).Value
    End If
    rs.Close
    Set rs = Nothing
End Function

' Ejecuta una sentencia que no devuelve filas (INSERT/UPDATE/DELETE).
Sub Execute(sql)
    Dim cn
    Set cn = GetConn()
    cn.Execute sql
    CloseConn(cn)
End Sub

' Ejecuta un INSERT y devuelve el IDENTITY generado.
Function ExecuteInsert(sql)
    Dim cn, rs
    Set cn = GetConn()
    cn.Execute sql
    Set rs = cn.Execute("SELECT CAST(SCOPE_IDENTITY() AS int) AS NUEVO_ID")
    If rs.EOF Then
        ExecuteInsert = 0
    Else
        ExecuteInsert = CLng(rs.Fields("NUEVO_ID").Value)
    End If
    rs.Close
    CloseConn(cn)
End Function

' Ejecuta varias sentencias dentro de una transaccion.
Sub ExecuteTx(sqlArray)
    Dim cn, i
    Set cn = GetConn()
    On Error Resume Next
    cn.BeginTrans
    For i = 0 To UBound(sqlArray)
        cn.Execute sqlArray(i)
        If Err.Number <> 0 Then Exit For
    Next
    If Err.Number <> 0 Then
        cn.RollbackTrans
        CloseConn(cn)
        Err.Raise vbObjectError + 1, "ExecuteTx", "Error en la transaccion: " & Err.Description
    Else
        cn.CommitTrans
        CloseConn(cn)
    End If
    Err.Clear
    On Error GoTo 0
End Sub

%>
