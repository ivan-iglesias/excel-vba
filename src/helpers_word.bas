Option Explicit

Const TIME_TO_WAIT As Long = 500

'
' Align all tables from the given word active document.
'
' pPositon: 1   wdAlignRowCenter (default)
'           0   wdAlignRowLeft
'           2   wdAlignRowRight
'
Public Function AlignTable(ByVal pWord As Object, _
                           Optional ByVal pPositon As Long = 1) As Boolean

    If pPositon < 0 Or pPositon > 2 Then Exit Function

    Dim table As Object

    For Each table In pWord.ActiveDocument.Tables
        table.Rows.Alignment = pPositon
    Next

    AlignTable = True
End Function

'
' Embed a file in Excel file.
'
' pSheet: Sheet where I want to embed the excell file
' pFile: File to embed
' pLine: Line to place the Excel file icon
' pColumn: Column to place the Excel file icon
' pOffsetY (optional): Offset icon by Y pixels
' pOffsetX (optional): Offset icon by X pixels
' pIcon: excel.exe, wordicon.exe
'
Public Sub EmbedFileInExcel(ByVal pSheet As Object, _
                            ByVal pFile As String, _
                            ByVal pLine As Long, _
                            ByVal pColumn As Variant, _
                            Optional ByVal pOffsetY As Long = 0, _
                            Optional ByVal pOffsetX As Long = 5, _
                            Optional ByVal pIcon As String = "excel.exe")

    pSheet.OLEObjects.add FileName:=pFile, _
                         Link:=False, _
                         DisplayAsIcon:=True, _
                         IconFileName:="excel.exe", _
                         IconIndex:=0, _
                         IconLabel:=FileName(pFile)

    pSheet.Shapes(pSheet.Shapes.Count).Width = 75
    pSheet.Shapes(pSheet.Shapes.Count).Top = pSheet.Cells(pLine, pColumn).Offset.Top + pOffsetY
    pSheet.Shapes(pSheet.Shapes.Count).Left = pSheet.Cells(pLine, pColumn).Offset.Left + pOffsetX
End Function


'
' Embed a file in word file.
'
' pWord: Word application object
' pTag:
' pFile
' pIcon: excel.exe, wordicon.exe
'
Public Sub EmbedFileInWord(ByVal pWord As Object, _
                           ByVal pTag As String, _
                           ByVal pFile As String, _
                           Optional ByVal pIcon As String = "excel.exe")

    With pWord.Selection.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .wrap = 1 ' wdFindContinue
        .Text = pTag
    End With

    If pWord.Selection.Find.Execute Then
        pWord.Selection.InlineShapes.AddOLEObject _
            FileName:=pFile, _
            LinkToFile:=False, _
            DisplayAsIcon:=True, _
            IconFileName:=pIcon, _
            IconLabel:=FileName(pFile)

        Sleep TIME_TO_WAIT
        Call ReplaceTag(pWord.ActiveDocument, pTag, "")
    End If
End Sub



'
' Inserto una tabla excel al documento word.
'
Public Sub InsertExcelTable(ByVal pWord As Object, _
                            ByVal pSheet As Worksheet, _
                            Optional ByVal pMainColumn As Variant = "A")

    Dim tag As String: tag = pSheet.Name

    ' Busco la etiqueta en funcion del nombre de la pestaña
    With pWord.Selection.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .wrap = 1 ' wdFindContinue
        .Text = tag
    End With

    If pWord.Selection.Find.Execute Then
        ' Si en la primera celda de la pestaña pone ELIMINAR es que
        ' la tabla no existe por lo que eliminamos la etiqueta.
        If pSheet.Cells(1, "A") = "ELIMINAR" Then
            Call ReplaceTag(pWord.ActiveDocument, tag, "")
            Exit Sub
        End If

        ' Obtengo el tamaño de la tabla
        Dim lastColumn As String
        Dim lastRow As Long

        lastColumn = NumberToLetter(pSheet.Cells(1, pSheet.Columns.Count).End(xlToLeft).Column)
        lastRow = pSheet.Cells(pSheet.Rows.Count, pMainColumn).End(xlUp).Row

        If lastColumn = pMainColumn And lastRow = 1 Then Exit Sub

        ' Selecciono el rango de celdas a apegar
        pSheet.Range("A1:" & lastColumn & lastRow).Copy

        ' Pego el rango en el word
        pWord.Selection.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False

        ' Dejo un tiempo para que termine y no de problemas
        Sleep TIME_TO_WAIT
    End If
End Sub


'
' Reemplazo una etiqueta por una imagen en un documento word.
'
Public Sub InsertImage(ByVal pWord As Object, _
                       ByVal pTag As String, _
                       Optional ByVal pPictureFile As String = "")

    ' Busco la etiqueta en funcion del nombre de la pestaña
    With pWord.Selection.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .wrap = 1 ' wdFindContinue
        .Text = pTag
    End With

    If pWord.Selection.Find.Execute Then
        ' Si no se pasa una imagen borramos la etiqueta del fichero.
        If pPictureFile = "" Then
            Call ReplaceTag(pWord.ActiveDocument, pTag, "")
            Exit Sub
        End If

        pWord.Selection.InlineShapes.AddPicture FileName:=pPictureFile, LinkToFile:=False, SavewithDocument:=True ', Range:=Selection.Range
        Sleep TIME_TO_WAIT
    End If
End Sub


'
' Reemplazo una etiqueta por una texto en un documento word.
'
Public Sub ReplaceTag(ByVal pDoc As Object, _
                      ByVal pTag As String, _
                      ByVal pValue As String)

    With pDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        .Execute FindText:=pTag, _
                 ReplaceWith:=pValue, _
                 Replace:=2 ' wdReplaceAll
    End With
End Sub
