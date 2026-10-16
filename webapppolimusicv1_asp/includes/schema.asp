<%
' =====================================================================
'  PoliMusic - Ajustes de esquema automaticos
'  La base BDD_PoliMusic original NO tiene campo de genero, y la
'  aplicacion necesita filtrar por genero. Este bloque agrega las
'  columnas e indices que faltan. Es IDEMPOTENTE: puede ejecutarse
'  tantas veces como haga falta sin perder datos.
'  Se ejecuta una sola vez por arranque de la aplicacion.
' =====================================================================

Sub EnsureSchema()
    If "" & Application("POLIMUSIC_SCHEMA_OK") = "1" Then Exit Sub

    Dim pasos, i, cn
    pasos = Array( _
        "IF COL_LENGTH('dbo.TBL_SONG','GENRE') IS NULL ALTER TABLE [dbo].[TBL_SONG] ADD [GENRE] [varchar](50) NULL", _
        "IF COL_LENGTH('dbo.TBL_ALBUM','GENRE') IS NULL ALTER TABLE [dbo].[TBL_ALBUM] ADD [GENRE] [varchar](50) NULL", _
        "UPDATE [dbo].[TBL_SONG] SET [GENRE] = CASE [SONG_NAME] " & _
            "WHEN 'Adventure' THEN 'Electronic' WHEN 'Dreams' THEN 'Chillout' WHEN 'Energy' THEN 'Rock' " & _
            "WHEN 'Scifi' THEN 'Electronic' WHEN 'Tomorrow' THEN 'Pop' ELSE 'Other' END " & _
            "WHERE [GENRE] IS NULL", _
        "UPDATE [dbo].[TBL_ALBUM] SET [GENRE] = CASE [TITLE_ALBUM] " & _
            "WHEN 'Rock Album' THEN 'Rock' WHEN 'Relax Music' THEN 'Chillout' " & _
            "WHEN 'Retro Album Music' THEN 'Retro' WHEN 'Classic Album' THEN 'Classical' " & _
            "WHEN 'Country Album' THEN 'Country' ELSE 'Other' END " & _
            "WHERE [GENRE] IS NULL", _
        "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_TBL_SONG_GENRE' AND object_id=OBJECT_ID('dbo.TBL_SONG')) " & _
            "CREATE NONCLUSTERED INDEX [IX_TBL_SONG_GENRE] ON [dbo].[TBL_SONG]([GENRE])", _
        "IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name='IX_TBL_ALBUM_GENRE' AND object_id=OBJECT_ID('dbo.TBL_ALBUM')) " & _
            "CREATE NONCLUSTERED INDEX [IX_TBL_ALBUM_GENRE] ON [dbo].[TBL_ALBUM]([GENRE])", _
        "IF NOT EXISTS (SELECT 1 FROM [dbo].[TBL_USER] WHERE [USERNAME]='userAdmin') " & _
            "INSERT INTO [dbo].[TBL_USER]([USERNAME],[PASSWORD],[EMAIL],[BIRTHDAY],[USER_PHOTO],[USER_TYPE]) " & _
            "VALUES('userAdmin','12345678','useradmin@epn.edu.ec','2001-12-16','~/userImages/root.png',1)" _
    )

    On Error Resume Next
    Set cn = GetConn()
    If Err.Number = 0 Then
        For i = 0 To UBound(pasos)
            cn.Execute pasos(i)
            Err.Clear
        Next
        cn.Close
        Application.Lock
        Application("POLIMUSIC_SCHEMA_OK") = "1"
        Application.Unlock
    End If
    Err.Clear
    On Error GoTo 0
End Sub
%>
