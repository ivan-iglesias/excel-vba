Option Explicit

Public Const TIME_TO_SLEEP As Long = 300

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

'
' Align all tables and images from the given word active document.
'
' pWord: Word application object
' pAlignImages (optional): Align document's images
' pPositon: (optional) 1   wdAlignRowCenter (default)
'                      0   wdAlignRowLeft
'                      2   wdAlignRowRight
'
Public Function AlignItems(ByVal pWord As Object, _
                           Optional ByVal pAlignImages As Boolean = False, _
                           Optional ByVal pPositon As Long = 1) As Boolean

    If pPositon < 0 Or pPositon > 2 Then Exit Function

    Dim item As Object

    For Each item In pWord.ActiveDocument.Tables
        item.Rows.Alignment = pPositon
    Next

    If pAlignImages Then
        For Each item In pWord.ActiveDocument.InlineShapes
            item.Range.ParagraphFormat.Alignment = pPositon
        Next
    End If

    AlignItems = True
End Function

'
' Embed a file in an Excel file.
'
' pSheet: Sheet where I want to embed the excell file
' pFile: File to embed
' pLine: Line to place the Excel file icon
' pColumn: Column to place the Excel file icon
' pOffsetY (optional): Offset icon by Y pixels
' pOffsetX (optional): Offset icon by X pixels
' pIcon (optional): excel.exe, wordicon.exe
' pIconIndex (optional): Icon's index
'
Public Function EmbedFileInExcel(ByVal pSheet As Object, _
                                 ByVal pFile As String, _
                                 ByVal pLine As Long, _
                                 ByVal pColumn As Variant, _
                                 Optional ByVal pOffsetY As Long = 0, _
                                 Optional ByVal pOffsetX As Long = 5, _
                                 Optional ByVal pIcon As String = "excel.exe", _
                                 Optional ByVal pIconIndex As Long = 0) As Boolean

    On Error GoTo errHandler

    pSheet.OLEObjects.Add fileName:=pFile, _
                         Link:=False, _
                         DisplayAsIcon:=True, _
                         IconFileName:=pIcon, _
                         IconIndex:=pIconIndex, _
                         IconLabel:=GetFileName(pFile)

    pSheet.Shapes(pSheet.Shapes.Count).Width = 75
    pSheet.Shapes(pSheet.Shapes.Count).Top = pSheet.Cells(pLine, pColumn).Offset.Top + pOffsetY
    pSheet.Shapes(pSheet.Shapes.Count).Left = pSheet.Cells(pLine, pColumn).Offset.Left + pOffsetX

    EmbedFileInExcel = True
errHandler:
End Function

'
' Embed a file in a word file.
'
' pWord: Word application object
' pTag: Tag name where to place the file
' pFile: File to embed
' pIconFileName (optional): excel.exe, wordicon.exe
'
Public Function EmbedFileInWord(ByVal pWord As Object, _
                                ByVal pTag As String, _
                                ByVal pFile As String, _
                                Optional ByVal pIconFileName As String = "excel.exe") As Boolean

On Error GoTo errHandler

    With pWord.Selection.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .wrap = 1 ' wdFindContinue
        .Text = pTag
    End With

    If pWord.Selection.Find.Execute Then

        pWord.Selection.InlineShapes.AddOLEObject _
            fileName:=pFile, _
            LinkToFile:=False, _
            DisplayAsIcon:=True, _
            IconFileName:=pIconFileName, _
            IconLabel:=GetFileName(pFile)

        Sleep TIME_TO_SLEEP
        Call ReplaceTag(pWord.ActiveDocument, pTag, "")
    End If

    EmbedFileInWord = True
errHandler:
End Function

'
' Insert an excel table into word document. If the first cell value is equal
' to 'delete', means that table does not exits, so we just delete the tag.
'
' pTag: Tag name where to place the file
' pSheet: Excel's sheet with table
' pWord: Word application object
' pPasteAsImage (optional): Paste table as image
' pLastLine (optional): Last table line, to add a fixed size table
' pLastColumn (optional): excel.exe, wordicon.exe
'
Public Function InsertExcel(ByVal pTag As String, _
                            ByVal pSheet As Worksheet, _
                            ByVal pWord As Object, _
                            Optional pPasteAsImage As Boolean = False, _
                            Optional pLastLine As Long = 0, _
                            Optional pLastColumn As String = "") As Boolean

On Error GoTo errHandler

    With pWord.Selection.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .wrap = 1 ' wdFindContinue
        .Text = pTag
    End With

    If pWord.Selection.Find.Execute Then
        If TrimLower(pSheet.Cells(1, "A")) = "delete" Then
            Call ReplaceTag(pWord.ActiveDocument, pTag, "")
            InsertExcel = True
            Exit Function
        End If

        If pLastLine = 0 Then
            pLastLine = pSheet.UsedRange.Rows(pSheet.UsedRange.Rows.Count).Row
        End If

        If pLastColumn = "" Then
            pLastColumn = NumberToLetter(pSheet.UsedRange.Columns(pSheet.UsedRange.Columns.Count).Column)
        End If

        If pLastColumn = "A" And pLastLine = 1 Then
            InsertExcel = True
            Exit Function
        End If

        If pPasteAsImage Then
            pSheet.Range("A1:" & pLastColumn & pLastLine).CopyPicture xlScreen, xlPicture
            Sleep TIME_TO_SLEEP
            pWord.Selection.Paste
            Exit Function
        End If

        pSheet.Range("A1:" & pLastColumn & pLastLine).Copy
        Sleep TIME_TO_SLEEP
        pWord.Selection.PasteExcelTable LinkedToExcel:=False, WordFormatting:=False, RTF:=False
    End If

    InsertExcel = True
errHandler:
End Function

'
' Insert an image into word document.
'
' pWord: Word application object
' pTag: Tag name where to place the image
' pPicture: Shape or images's full path to insert
'
Public Function InsertImage(ByVal pWord As Object, _
                            ByVal pTag As String, _
                            ByVal pPicture As Variant) As Boolean

On Error GoTo errHandler

    With pWord.Selection.Find
        .ClearFormatting
        .MatchCase = False
        .MatchWholeWord = True
        .wrap = 1 ' wdFindContinue
        .Text = pTag
    End With

    If pWord.Selection.Find.Execute Then

        If VarType(pPicture) = 8 Then
            ' String (image's full path)
            pWord.Selection.InlineShapes.AddPicture fileName:=pPicture, LinkToFile:=False, SavewithDocument:=True ', Range:=Selection.Range

        ElseIf VarType(pPicture) = 9 Then
            ' Object (shape)
            pPicture.Copy
            Sleep TIME_TO_SLEEP

            ' 3 wdPasteMetafilePicture
            ' 0 wdInLine
            pWord.Selection.PasteSpecial Link:=False, DataType:=3, Placement:=0, DisplayAsIcon:=False
        End If
    End If

    InsertImage = True
errHandler:
End Function

'
' Replace a text from word document.
'
' pDoc: Active document
' pTag: Tag name to be replaced
' pValue: New text
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

        .Execute FindText:=pTag, ReplaceWith:=pValue, Replace:=2 ' wdReplaceAll
    End With
End Sub
