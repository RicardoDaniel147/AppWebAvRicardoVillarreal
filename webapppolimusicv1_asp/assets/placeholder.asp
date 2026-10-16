<%@ Language="VBScript" %>
<%
' =====================================================================
'  PoliMusic - Generador de imagenes SVG de relleno
'  Se usa cuando falta la portada de un album, la foto de un artista
'  o el avatar de un usuario, para que nunca se vean imagenes rotas.
' =====================================================================
Option Explicit

Dim kind, seed, h, paleta, desde, hasta, etiqueta, icono

kind = LCase(Trim("" & Request.QueryString("kind")))
seed = Trim("" & Request.QueryString("seed"))
If kind = "" Then kind = "album"
If seed = "" Then seed = "PoliMusic"

Response.ContentType = "image/svg+xml"
Response.Charset = "utf-8"
Response.AddHeader "Cache-Control", "public, max-age=86400"
Response.Expires = 1440

h = HashSeed(seed)
paleta = h Mod 8

Select Case paleta
    Case 0 : desde = "#0B1E3B" : hasta = "#1E5BFF"
    Case 1 : desde = "#081226" : hasta = "#00C2FF"
    Case 2 : desde = "#160B2E" : hasta = "#7A5CFF"
    Case 3 : desde = "#06232A" : hasta = "#00E0B8"
    Case 4 : desde = "#1A0B22" : hasta = "#FF4FD8"
    Case 5 : desde = "#0A1F2E" : hasta = "#2F7CFF"
    Case 6 : desde = "#101425" : hasta = "#4EA1FF"
    Case Else : desde = "#0D1A16" : hasta = "#38F2A0"
End Select

etiqueta = Iniciales(seed, kind)
icono = IconoSvg(kind)

Response.Write "<svg xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 400 400"" width=""400"" height=""400"" role=""img"" aria-label=""" & Esc(seed) & """>" & vbCrLf
Response.Write "  <defs>" & vbCrLf
Response.Write "    <linearGradient id=""g"" x1=""0"" y1=""0"" x2=""1"" y2=""1"">" & vbCrLf
Response.Write "      <stop offset=""0%"" stop-color=""" & desde & """/>" & vbCrLf
Response.Write "      <stop offset=""100%"" stop-color=""" & hasta & """/>" & vbCrLf
Response.Write "    </linearGradient>" & vbCrLf
Response.Write "    <radialGradient id=""glow"" cx=""0.25"" cy=""0.2"" r=""0.85"">" & vbCrLf
Response.Write "      <stop offset=""0%"" stop-color=""#FFFFFF"" stop-opacity=""0.30""/>" & vbCrLf
Response.Write "      <stop offset=""100%"" stop-color=""#FFFFFF"" stop-opacity=""0""/>" & vbCrLf
Response.Write "    </radialGradient>" & vbCrLf
Response.Write "  </defs>" & vbCrLf
Response.Write "  <rect width=""400"" height=""400"" fill=""url(#g)""/>" & vbCrLf
Response.Write "  <rect width=""400"" height=""400"" fill=""url(#glow)""/>" & vbCrLf
Response.Write "  <g opacity=""0.16"" stroke=""#FFFFFF"" stroke-width=""1"">" & vbCrLf
Response.Write RejillaSvg()
Response.Write "  </g>" & vbCrLf
Response.Write "  <g fill=""none"" stroke=""#FFFFFF"" stroke-opacity=""0.55"" stroke-width=""9"" stroke-linecap=""round"" stroke-linejoin=""round"">" & vbCrLf
Response.Write "    " & icono & vbCrLf
Response.Write "  </g>" & vbCrLf
Response.Write "  <text x=""200"" y=""330"" font-family=""Segoe UI, Inter, Arial, sans-serif"" font-size=""58"" font-weight=""700"" fill=""#FFFFFF"" fill-opacity=""0.92"" text-anchor=""middle"" letter-spacing=""4"">" & Esc(etiqueta) & "</text>" & vbCrLf
Response.Write "</svg>" & vbCrLf

' ---------------------------------------------------------------- funciones
Function HashSeed(s)
    Dim i, c, hh
    hh = 7
    For i = 1 To Len(s)
        c = Asc(Mid(s, i, 1))
        hh = ((hh * 31) + c) Mod 1000003
    Next
    HashSeed = hh
End Function

Function Iniciales(s, k)
    Dim partes, i, r
    If k = "artist" Or k = "user" Then
        partes = Split(Trim(s), " ")
        If UBound(partes) >= 1 Then
            Iniciales = UCase(Left(Trim(partes(0)), 1)) & UCase(Left(Trim(partes(1)), 1))
            Exit Function
        End If
    End If
    s = Trim(s)
    If Len(s) = 0 Then
        Iniciales = "PM"
    ElseIf Len(s) = 1 Then
        Iniciales = UCase(s)
    Else
        Iniciales = UCase(Left(s, 2))
    End If
End Function

Function IconoSvg(k)
    Select Case k
        Case "artist", "user"
            IconoSvg = "<circle cx='200' cy='150' r='52'/><path d='M104 300c0-53 43-96 96-96s96 43 96 96'/>"
        Case "playlist"
            IconoSvg = "<path d='M132 148h136M132 200h136M132 252h88'/><circle cx='288' cy='276' r='26' fill='#FFFFFF' fill-opacity='0.55' stroke='none'/>"
        Case Else
            IconoSvg = "<circle cx='200' cy='200' r='96'/><circle cx='200' cy='200' r='30'/><path d='M296 200a96 96 0 0 1-96 96'/>"
    End Select
End Function

Function RejillaSvg()
    Dim s, x, y
    s = ""
    For x = 40 To 360 Step 40
        s = s & "<line x1='" & x & "' y1='0' x2='" & x & "' y2='400'/>"
    Next
    For y = 40 To 360 Step 40
        s = s & "<line x1='0' y1='" & y & "' x2='400' y2='" & y & "'/>"
    Next
    RejillaSvg = s
End Function

Function Esc(v)
    Dim s
    s = "" & v
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", " ")
    s = Replace(s, "'", " ")
    Esc = s
End Function
%>
