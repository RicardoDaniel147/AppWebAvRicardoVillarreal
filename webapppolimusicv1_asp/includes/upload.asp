<%
' =====================================================================
'  PoliMusic - Subida de archivos (multipart/form-data)
'  Implementado en VBScript puro, SIN componentes externos
'  (no requiere Persits.Upload ni similares).
' =====================================================================

Dim gFields, gFiles, gFileNames, gUploadError

' Comprueba si la peticion actual es multipart.
Function IsMultipart()
    IsMultipart = (InStr(LCase("" & Request.ServerVariables("CONTENT_TYPE")), "multipart/form-data") > 0)
End Function

' ---------------------------------------------------- conversion binaria
' Convierte un array de bytes en una cadena 1:1 (latin-1).
Function BinToLatin1(bin)
    Dim st
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write bin
    st.Position = 0
    st.Type = 2
    st.Charset = "iso-8859-1"
    BinToLatin1 = st.ReadText
    st.Close
    Set st = Nothing
End Function

' Convierte una cadena 1:1 (latin-1) de vuelta a bytes.
Function Latin1ToBin(s)
    Dim st
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "iso-8859-1"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    Latin1ToBin = st.Read
    st.Close
    Set st = Nothing
End Function

' Interpreta una cadena 1:1 como UTF-8 y la decodifica.
Function Utf8FromLatin1(s)
    Dim st, bin
    On Error Resume Next
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "iso-8859-1"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    bin = st.Read
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    Utf8FromLatin1 = st.ReadText
    st.Close
    Set st = Nothing
    If Err.Number <> 0 Then
        Utf8FromLatin1 = s
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------ parser
' Extrae el boundary del Content-Type.
Function ExtractBoundary(ct)
    Dim p, s, i
    ExtractBoundary = ""
    p = InStr(LCase(ct), "boundary=")
    If p = 0 Then Exit Function
    s = Mid(ct, p + 9)
    If Left(s, 1) = """" Then
        s = Mid(s, 2)
        i = InStr(s, """")
        If i > 0 Then s = Left(s, i - 1)
    Else
        i = InStr(s, ";")
        If i > 0 Then s = Left(s, i - 1)
    End If
    ExtractBoundary = Trim(s)
End Function

' Lee el valor de un atributo del encabezado:  name="valor"
Function AttrValue(line, attr)
    Dim re, m
    AttrValue = ""
    On Error Resume Next
    Set re = New RegExp
    re.Pattern = attr & "=""([^""]*)"""
    re.IgnoreCase = True
    Set m = re.Execute(line)
    If m.Count > 0 Then AttrValue = m(0).SubMatches(0)
    Set re = Nothing
    Err.Clear
    On Error GoTo 0
End Function

' Carpeta temporal propia (dentro de la aplicacion) para los archivos subidos.
Function UploadTempDir()
    Dim p
    p = MediaPhysicalRoot() & "\uploads"
    If Not GetFSO().FolderExists(p) Then GetFSO().CreateFolder p
    UploadTempDir = p
End Function

' Guarda un cuerpo binario (cadena 1:1) en un archivo temporal y devuelve su ruta.
Function SaveTempBinary(s)
    Dim st, path
    path = UploadTempDir() & "\pm_" & Year(Now) & Right("0" & Month(Now), 2) & _
           Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & _
           Right("0" & Second(Now), 2) & CStr(Int(Rnd() * 100000)) & ".tmp"
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write Latin1ToBin(s)
    st.SaveToFile path, 2
    st.Close
    Set st = Nothing
    SaveTempBinary = path
End Function

' Procesa cada parte del formulario multipart.
Sub ParsePart(part)
    Dim p, sep, headers, body, name, fileName, lines, j, line, tmpPath

    If Left(part, 2) = vbCrLf Then part = Mid(part, 3)
    If Left(part, 2) = "--" Then Exit Sub
    If Len(part) = 0 Then Exit Sub

    sep = InStr(part, vbCrLf & vbCrLf)
    If sep = 0 Then Exit Sub

    headers = Left(part, sep - 1)
    body = Mid(part, sep + 4)
    If Right(body, 2) = vbCrLf Then body = Left(body, Len(body) - 2)

    name = ""
    fileName = ""
    lines = Split(headers, vbCrLf)
    For j = 0 To UBound(lines)
        line = lines(j)
        If InStr(LCase(line), "content-disposition:") = 1 Then
            name = AttrValue(line, "name")
            fileName = AttrValue(line, "filename")
        End If
    Next

    If name = "" Then Exit Sub

    If fileName <> "" Then
        tmpPath = SaveTempBinary(body)
        gFiles(name) = tmpPath
        gFileNames(name) = fileName
    Else
        gFields(name) = Utf8FromLatin1(body)
    End If
End Sub

' Punto de entrada: analiza la peticion multipart completa.
Sub ParseMultipart()
    Dim ct, boundary, totalBytes, rawBin, rawStr, delim, partes, i

    Set gFields = Server.CreateObject("Scripting.Dictionary")
    Set gFiles = Server.CreateObject("Scripting.Dictionary")
    Set gFileNames = Server.CreateObject("Scripting.Dictionary")
    gUploadError = ""

    If Not IsMultipart() Then Exit Sub

    ct = "" & Request.ServerVariables("CONTENT_TYPE")
    boundary = ExtractBoundary(ct)
    If boundary = "" Then
        gUploadError = "No se pudo determinar el limite del formulario."
        Exit Sub
    End If

    totalBytes = Request.TotalBytes
    If totalBytes <= 0 Then Exit Sub

    On Error Resume Next
    rawBin = Request.BinaryRead(totalBytes)
    If Err.Number <> 0 Then
        gUploadError = "No se pudo leer el cuerpo de la peticion (" & Err.Description & ")."
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    rawStr = BinToLatin1(rawBin)
    delim = "--" & boundary
    partes = Split(rawStr, delim)

    For i = 1 To UBound(partes)
        ParsePart partes(i)
    Next
End Sub

' --------------------------------------------------------- acceso a los datos
' Valor de un campo del formulario (funciona con multipart y con urlencoded).
Function PostVal(name)
    If IsMultipart() Then
        If IsObject(gFields) Then
            If gFields.Exists(name) Then
                PostVal = Trim("" & gFields(name))
                Exit Function
            End If
        End If
        PostVal = ""
        Exit Function
    End If
    PostVal = Trim("" & Request.Form(name))
End Function

' Indica si se subio un archivo en el campo indicado.
Function HasUpload(fieldName)
    HasUpload = False
    If Not IsObject(gFiles) Then Exit Function
    If gFiles.Exists(fieldName) Then
        HasUpload = (Len("" & gFiles(fieldName)) > 0)
    End If
End Function

Function UploadOriginalName(fieldName)
    UploadOriginalName = ""
    If Not IsObject(gFileNames) Then Exit Function
    If gFileNames.Exists(fieldName) Then UploadOriginalName = "" & gFileNames(fieldName)
End Function

Function ExtOf(fileName)
    Dim i
    ExtOf = ""
    i = InStrRev(fileName, ".")
    If i > 0 Then ExtOf = LCase(Mid(fileName, i))
End Function

' Mueve el archivo subido a la carpeta de recursos y devuelve la ruta para la BD.
' allowedExt es una lista tipo "|.jpg|.png|.gif|".
Function SaveUpload(fieldName, destFolder, allowedExt)
    Dim tmp, original, ext, nuevo, destPath, sello

    SaveUpload = ""
    If Not HasUpload(fieldName) Then Exit Function

    tmp = "" & gFiles(fieldName)
    If Not FileExists(tmp) Then Exit Function

    original = UploadOriginalName(fieldName)
    ext = ExtOf(original)
    If ext = "" Then Exit Function
    If InStr(allowedExt, "|" & ext & "|") = 0 Then Exit Function

    sello = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
            Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2) & _
            CStr(Int(Rnd() * 100000))
    nuevo = sello & ext

    destPath = MediaPhysicalRoot() & "\" & destFolder & "\" & nuevo
    On Error Resume Next
    GetFSO().CopyFile tmp, destPath, True
    If Err.Number <> 0 Then
        gUploadError = "No se pudo guardar el archivo: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' El archivo temporal ya no hace falta.
    On Error Resume Next
    GetFSO().DeleteFile tmp, True
    Err.Clear
    On Error GoTo 0

    SaveUpload = "~/" & destFolder & "/" & nuevo
End Function

' Borra los archivos temporales que hayan quedado de peticiones anteriores.
Sub CleanupUploadTemp()
    Dim fso, fld, f, limite
    On Error Resume Next
    Set fso = GetFSO()
    If fso.FolderExists(UploadTempDir()) Then
        Set fld = fso.GetFolder(UploadTempDir())
        limite = Now() - 0.5 ' 12 horas
        For Each f In fld.Files
            If f.DateLastModified < limite Then fso.DeleteFile f.Path, True
        Next
    End If
    Err.Clear
    On Error GoTo 0
End Sub

%>
