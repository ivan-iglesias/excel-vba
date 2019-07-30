Option Explicit


'
' Create an array of a given length.
'
' pLength: Array length
' pValue: Value to initialize the array
'
' RETURN: Array
'
Public Function ArrayInitialize(ByVal pLength, _
                                ByVal pValue As Variant) As Variant

    Dim myArray(0 To pLength) As Variant
    Dim i As Long

    For i = 0 To pLength
        myArray(i) = Array(i + 1, pValue)
    Next

    ArrayInitialize = myArray
End Function


'
' Join strings from a collection.
'
' pCollection: Collection of strings
' pDelimiter: Delimiter between strings
'
' RETURN: String
'
Public Function CollectionToString(ByVal pCollection As Collection, _
                                   ByVal pDelimiter As String) As String

    Dim item As Variant

    For Each item In pCollection
        CollectionToString = CollectionToString & item & pDelimiter
    Next

    If Len(CollectionToString) > 0 Then
        CollectionToString = Left(CollectionToString, Len(CollectionToString) - Len(pDelimiter))
    End If
End Function


'
' Join two collections.
'
' collectionA: Collection to update with items from second collection.
' collectionB: Collection to add in the first collection.
'
Public Sub CollectionMerge(ByRef collectionA As Collection, _
                           ByVal collectionB As Collection)

    Dim i As Long

    For i = 1 To collectionB.Count
        collectionA.add collectionB.item(i)
    Next i
End Sub


'
' Return a specific day of the given year-week.
' DayNumber options: 1-Monday, 2-Tuesday, ..., 7-Sunday
'
' RETURN: Date
'
Public Function DayFromWeek(ByVal pYear As Long, _
                            ByVal pWeek As Long, _
                            ByVal pDayNumber As Long) As Date

    If pDayNumber < 1 Or pDayNumber > 7 Then
        Throw ("Not a valid day number, It must be between 1 and 7")
    End If

    Dim temp As Date: temp = DateSerial(pYear, 1, 1)

    temp = temp + (pDayNumber - Weekday(temp, vbMonday))

    DayFromWeek = DateAdd("ww", pWeek - 1, temp)
End Function


'
' Draw a border to the given range.
'
' pRange: Range in which to draw the border
'
Public Sub DrawRangeBorder(ByVal pRange As Range)
    With pRange
        .Borders.LineStyle = xlContinuous
        .Borders.Color = vbBlack
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThin
    End With
End Sub


'
' Enable-disable excel options to improve the performance.
'
' pIsOn (optional): True: Enable performance improve
'                   False: Disable performance improve
'
Public Sub EnableOptimization(Optional ByVal pIsOn As Boolean = True)
    Application.ScreenUpdating = Not pIsOn
    Application.DisplayAlerts = Not pIsOn
    Application.EnableEvents = Not pIsOn
    Application.Calculation = IIf(pIsOn, xlCalculationManual, xlCalculationAutomatic)
    Application.CutCopyMode = False
End Sub


'
' End the current process. It usually happens when I don't select a required file.
' Close the open books without saving the changes.
'
' pNumberBooks (optional): Number of books to close
'
Public Sub EndProcess(Optional ByVal pNumberBooks As Long = 0)
    If pNumberBooks > 0 Then
        Call ExcelClose(pNumberBooks, False)
    End If

    Call EnableOptimization(False)
    End
End Sub


'
' Close the last 'pNumberBooks' workbooks.
'
' pNumberBooks: Number of books to close
' pSave: True: Save changes
'        False: Not save changes
'
' RETURN: True  > OK
'         False > NOK
'
Public Function ExcelClose(ByRef pNumberBooks As Long, _
                           Optional ByVal pSave As Boolean = False) As Boolean

    If pNumberBooks < 1 Or pNumberBooks > Workbooks.Count Then Exit Function

    On Error GoTo errHandler

    Dim i As Long
    Dim n As Long

    For i = Workbooks.Count To Workbooks.Count - pNumberBooks + 1 Step -1
        Workbooks(i).Close savechanges:=pSave
        n = n + 1
    Next

    pNumberBooks = pNumberBooks - n
    ExcelClose = True
    Exit Function
errHandler:
End Function


'
' Create a new excel file.
'
' pFile: Destination full path
' pNumberSheets: Number of sheets to add to the new file
'
' RETURN: True  > OK
'         False > NOK
'
Public Function ExcelCreate(ByVal pFile As String, _
                            Optional ByVal pNumberSheets As Long = 1) As Boolean

    On Error GoTo errHandler

    Dim idxExcel As Long: idxExcel = Workbooks.Count + 1

    If pNumberSheets > 0 Then
        Application.SheetsInNewWorkbook = pNumberSheets
        Application.Workbooks.add
        Workbooks(idxExcel).SaveAs FileName:=pFile
        ExcelCreate = True
    End If

    Exit Function
errHandler:
End Function


'
' Open the given text file.
'
' pFile: File to open
' pTab: True : Use tab delimiter
' pSemicolon: True : Use semicolon delimiter
' pComa: True : Use comma delimiter
' pSpace: True : Use space delimiter
'
' RETURN: Excel > Boolean
'         CSV   > Full path of the file
'         TXT   > Full path of the file
'
Public Function ExcelOpen(ByVal pFile As String, _
                          Optional ByRef pBooksOpen As Long = 0, _
                          Optional ByVal pTab As Boolean = False, _
                          Optional ByVal pSemicolon As Boolean = False, _
                          Optional ByVal pComa As Boolean = False, _
                          Optional ByVal pSpace As Boolean = False) As Variant

    On Error GoTo errHandler

    Dim extension As String: extension = LCase(FileExtension(pFile))

    ' Excel file
    If extension = "xls" Or extension = "xlsx" Then
        Workbooks.Open FileName:=pFile, UpdateLinks:=xlUpdateLinksAlways ', corruptload:=xlRepairFile
        pBooksOpen = pBooksOpen + 1
        ExcelOpen = True
        Exit Function
    End If

    ' CSV file
    If extension = "csv" Then
        pFile = Left(pFile, Len(pFile) - 3) & "txt"
        If FileExists(pFile) Then FileDelete (pFile)
        Call FileCopy(pFile, pFile)
    Else
        ExcelOpen = "-1"
        Exit Function
    End If

    ExcelOpen = pFile

    Dim myArray As Variant: Set myArray = ArrayInitialize(300, 2)

    Workbooks.OpenText FileName:=pFile, _
                       StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                       ConsecutiveDelimiter:=True, Tab:=pTab, Semicolon:=pSemicolon, Comma:=pComa, Space:=pSpace, Other:=False, _
                       fieldInfo:=myArray, Local:=True

    pBooksOpen = pBooksOpen + 1
    Exit Function
errHandler:
    ExcelOpen = "-1"
End Function


'
' Export an excell file to access
'
' pWorkbook: Excel file to export. Must be open.
' pAccessFile: Access full path (xxxxx.accdb)
'
Public Sub ExcelToAccess(ByVal pWorkbook As Workbook, _
                         ByVal pAccessFile As String)

    Const acImport = 0
    Const acSpreadsheetTypeExcel9 = 8

    Dim objAccess As Object
    Dim sheet As Variant

    Set objAccess = CreateObject("Access.Application")
    objAccess.NewCurrentDatabase pAccessFile

    For Each sheet In pWorkbook.Worksheets
        objAccess.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, sheet.Name, pWorkbook.FullName, True, sheet.Name & "$"
    Next

    Set objAccess = Nothing
End Sub


'
' Check if given key exists in a collection.
'
' pCollection: Collection
' pKey: Key
'
' RETURN: True  > Exists
'         False > Not exists
'
Public Function ExistsKey(ByRef pCollection As Collection, _
                          ByVal pKey As Variant) As Boolean

    On Error GoTo errHandler
    pCollection.item (pKey)
    ExistsKey = True
errHandler:
End Function


'
' Copy a file.
'
' pSource: Source file to copy
' pDestination: Destination full path
'
Public Sub FileCopy(ByVal pSource As String, _
                    ByVal pDestination As String)

    Dim xlobj As Object
    Set xlobj = CreateObject("Scripting.FileSystemObject")
    xlobj.CopyFile pSource, pDestination, True
    Set xlobj = Nothing
End Sub


'
' Create a file with the the given lines.
'
' pDestination: Destination full path
' pLines: Lines to print in a file
'
Public Sub FileCreate(ByVal pDestination As String, _
                      ByVal pLines As Collection)

    Dim line As Variant

    If pLines.Count > 0 Then
        Open pDestination For Output As #1
            For Each line In pLines
                Print #1, Tab(0); line
            Next
        Close #1
    End If
End Sub


'
' Delete a file.
'
' pFile: File's full path to delete
'
' RETURN: True  > OK
'         False > NOK
'
Public Function FileDelete(ByVal pFile As String) As Boolean

    On Error GoTo errHandler
    Kill pFile
    FileDelete = True
errHandler:
End Function


'
' Check if exists a file.
'
' pFile: File's full path to check
'
' RETURN: True  > Exists
'         False > Not exists
'
Public Function FileExists(ByVal pFile As String) As Boolean
    If Dir(pFile) <> "" Then FileExists = True
End Function


'
' Get the extension from file's full path.
'
' pFile: File full path
'
' RETURN: String
'
Public Function FileExtension(ByVal pFile As String) As String
    FileExtension = Right(pFile, Len(pFile) - InStrRev(pFile, "."))
End Function


'
' Get the file name from its full path.
'
' pFile: File full path
'
' RETURN: String
'
Public Function FileName(ByVal pFile As String) As String
    FileName = Right(pFile, Len(pFile) - InStrRev(pFile, Application.PathSeparator))
End Function


'
' Get the file path from its full path.
'
' pFile: File full path
'
' RETURN: String
'
Public Function FilePath(ByVal pFile As String) As String
    FilePath = Left(pFile, InStrRev(pFile, Application.PathSeparator))
End Function


'
' Create a new excel file.
'
' pPath: Folder's parent path
' pName: Folder's name
' pWithDateTime (Optional): Add timestamp to folder's name
'
' RETURN: String, OK  > Folder full path
'                 NOK > ""
'
Public Function FolderCreate(ByVal pPath As String, _
                             ByVal pName As String, _
                             Optional ByVal pWithDateTime As Boolean = True) As String

    On Error GoTo errHandler

    If Right(pPath, 1) <> Application.PathSeparator Then pPath = pPath & Application.PathSeparator

    If pWithDateTime Then pName = Format(Date, "YYYYMMDD") & "_" & Format(Time, "HHMM") & "_" & pName

    FolderCreate = pPath & pName & Application.PathSeparator

    If Dir(FolderCreate, vbDirectory) = "" Then MkDir (FolderCreate)

    Exit Function
errHandler:
    FolderCreate = ""
End Function


'
' Delete folder and it's content.
'
' pPath: Folder's path
'
Public Function FolderDelete(ByVal pPath As String) As Boolean
    On Error Resume Next
        If Right(pPath, 1) <> Application.PathSeparator Then pPath = pPath & Application.PathSeparator
        Kill pPath & "*.*"
        RmDir pPath
        FolderDelete = True
    On Error GoTo 0
End Function


'
' Get the files of a given folder.
'
' pPath: Folder's path
'
' RETURN: Collection of files
'
Public Function FolderFiles(ByVal pPath As String) As Collection
    Set FolderFiles = New Collection

    On Error GoTo errHandler

    Dim file As String: file = Dir(pPath)

    Do While file <> ""
        FolderFiles.add Directorio & file
        file = Dir()
    Loop
errHandler:
End Function


'
' Check if workbook is open.
'
' pFile: File's full path to check
'
' RETURN: True  > Open
'         False > Close
'
Function IsExcelOpen(ByVal pFile As String) As Boolean
    Dim FileName As String
    Dim wb As Variant

    On Error Resume Next

    FileName = LCase(FileName(pFile))

    For Each wb In Workbooks
        If LCase(wb.Name) = FileName Then
            ExcelIsOpen = True
            Exit Function
        End If
    Next wb
End Function


'
' Trim a string and convert it to lower case.
'
' pText: Text
'
' RETURN: Processed text
'
Public Function Normalize(ByVal pText As String, _
                          Optional ByVal pReplaceSpaceWithUnderscore As Boolean = False)

    Normalize = TrimLower(pText)

    Normalize = Replace(Replace(Normalize, Chr(10), ""), Chr(13), "")

    Do While InStr(Normalize, "  ") <> 0
        Normalize = Replace(Normalize, "  ", " ")
    Loop

    If pReplaceSpaceWithUnderscore Then Normalize = Replace(Normalize, " ", "_")
End Function


'
' Get the column letter for to the given number.
'
' pNumber: Number
'
' RETURN: Column letter
'
Public Function NumberToLetter(ByVal pNumber As Long) As String
   NumberToLetter = Replace(Cells(1, pNumber).Address(True, False), "$1", "")
End Function


'
' Returns a new string that right-aligns the characters in this instance by padding
' them on the left with a specified character, for a specified total length.
'
' pText: Text to apply the padding
' pLength: Text total length
' pChar: Padding character
'
' RETURN: String
'
Public Function PadLeft(ByVal pText As String, _
                        ByVal pLength As Long, _
                        ByVal pChar As String) As String

    If Len(pText) < pLength Then
        pText = String(pLength - Len(CStr(pText)), pChar) & CStr(pText)
    End If

    PadLeft = pText
End Function


'
' Returns a new string that left-aligns the characters in this instance by padding
' them on the right with a specified character, for a specified total length.
'
' pText: Text to apply the padding
' pLength: Text total length
' pChar: Padding character
'
' RETURN: String
'
Public Function PadRight(ByVal pText As Variant, _
                         ByVal pLength As Long, _
                         ByVal pChar As String) As String

    If Len(pText) < pLength Then
        pText = CStr(pText) & String(pLength - Len(CStr(pText)), pChar)
    End If

    PadLeft = pText
End Function


'
' Returns the column index position of a field.
'
' pSheet: Sheet where is the field to search
' pName: Text to search
' pRow (optional): Row where the text must be
'
' RETURN:  n > Column index position
'         -1 > Not Found
'
Public Function SearchColumn(ByVal pSheet As Worksheet, _
                             ByVal pName As String, _
                             Optional ByVal pRow As Long = 1, _
                             Optional ByRef pMessage As String = "na") As Long

    SearchColumn = -1

    Dim pNameNormalize As String: pNameNormalize = Normalize(pName)

    If pNameNormalize = "" Then Exit Function

    Dim pNumberColumns As Long
    Dim i As Long

    pNumberColumns = pSheet.UsedRange.Rows(pRow).Columns.Count - 1

    i = 1
    Do While Not IsEmpty(pSheet.Cells(pRow, i)) Or i <= pNumberColumns
        If Normalize(pSheet.Cells(pRow, i).value) = pNameNormalize Then
            SearchColumn = i
            Exit Function
        End If
        i = i + 1
    Loop

    If SearchColumn = -1 And pMessage <> "na" Then
        pMessage = IIf(pMessage = "", pName, pMessage & ", " & pName)
    End If
End Function


'
' Returns the column index position of a field.
'
' pSheet: Sheet where is the field to search
' pName: Text to search
' pColumn (optional): Column where the text must be
'
' RETURN:  n > Row index position
'         -1 > Not Found
'
Public Function SearchLine(ByVal pSheet As Worksheet, _
                           ByVal pName As String, _
                           Optional ByVal pColumn As Variant = 1) As Long

    SearchLine = -1

    pName = Normalize(pName)

    If pName = "" Then Exit Function

    Dim pNumberRows As Long
    Dim i As Long

    pNumberRows = pSheet.UsedRange.Columns(pColumn).Rows.Count - 1

    i = 1
    Do While Not IsEmpty(pSheet.Cells(i, pColumn)) Or i <= pNumberRows
        If Normalize(pSheet.Cells(i, pColumn).value) = pName Then
            SearchLine = i
            Exit Do
        End If
        i = i + 1
    Loop
End Function

'
' Search a shape in the given worksheet.
'
' pSheet: Worksheet
' pName: Shape's name
'
' RETURN: Shape
'
Public Function SearchShape(ByVal pSheet As Worksheet, _
                            ByVal pName As String) As Shape
    On Error Resume Next
    Set SearchShape = pSheet.Shapes(pName)
    Err.Clear
End Function

'
' Search and initialize a sheet by given sheet name in the workbook.
'
' pWorkbook: Workbook object
' pSheetName: Sheet's name
' pInitialize (optional): Initialize the sheet
'
' RETURN: Worksheet
'
Public Function SearchSheet(ByVal pWorkbook As Workbook, _
                            ByVal pSheetName As String, _
                            Optional ByVal pInitialize As Boolean = True) As Worksheet

    Dim idx As Long: idx = SheetIndexPosition(pWorkbook, pName)

    If idx = -1 Then Exit Function

    Set SearchSheet = pWorkbook.Sheets(idx)

    If pInitialize Then Call SheetInitialize(SearchSheet)
End Function


'
' Show file dialog to select a file.
'
' pHeader: File dialog header
' pType (optional): File's type (xls, csv, txt, ...)
' pFilter (optional): Text to filter the visible files in the file dialog
' pMultiSelect (optional): Select one or more files
'
' RETURN: Selected files:
'         - pMultipleFiles = True  > Collection
'         - pMultipleFiles = False > String
'
Public Function SelectFile(ByVal pHeader As String, _
                           Optional ByVal pType As String = "", _
                           Optional ByVal pFilter As String = "", _
                           Optional ByVal pMultiSelect As Boolean = False) As Variant

    ' Construct the filter text
    If pType <> "" And pFilter <> "" Then
        pFilter = "Filter (*" & pFilter & "*." & pType & "*), *" & pFilter & "*." & pType & "*"
    ElseIf pType <> "" Then
        pFilter = "Filter (*." & pType & "*), *." & pType & "*"
    ElseIf pFilter <> "" Then
        pFilter = "Filter (*" & pFilter & "*), *" & pFilter & "*"
    End If

    pFilter = IIf(pFilter = "", "", pFilter & ",") & "All Files (*.*),*.*"

    ' Open file dialog
    Dim files As Variant
    files = Application.GetOpenFilename(Title:=pHeader, _
                                        filefilter:=pFilter, _
                                        MultiSelect:=pMultiSelect)

    ' Process the result
    If pMultiSelect Then
        Set SelectFile = IIf(IsArray(files), ToCollection(files), New Collection)
        Exit Function
    End If

    SelectFile = IIf(files <> False, files, "")
End Function


'
' Show folder dialog to select a folder.
'
' pTitle (optional): Folder dialog header
'
' RETURN: Selected folder
'
Public Function SelectFolder(Optional ByVal pTitle As String = "") As String
    SelectFolder = ""

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = pTitle
        .AllowMultiSelect = False
        .Show
        On Error Resume Next
            SelectFolder = .SelectedItems(1)
            Err.Clear
        On Error GoTo 0
    End With
End Function


'
' Set a basic sheet style.
'
' pSheet: Sheet to apply the style
' pColumnWidth (optional): Columns width
' pRowHeight (optional): Rows height
'
Public Sub SheetBasicStyle(ByVal pSheet As Worksheet, _
                           Optional pColumnWidth As Long = -1, _
                           Optional pRowHeight As Long = -1)

    With pSheet
        .Cells.Font.Size = 11
        .Cells.Font.Name = "Calibri"
        .Columns().HorizontalAlignment = xlLeft
        .Columns().VerticalAlignment = xlCenter
        .Rows(1).HorizontalAlignment = xlCenter
    End With

    If pColumnWidth > -1 Then pSheet.Columns().ColumnWidth = pColumnWidth
    If pRowHeight > -1 Then pSheet.Rows().RowHeight = pRowHeight
End Sub


'
' Return the index position of a given sheet name.
'
' pWorkbook: Workbook object
' pSheetName: Sheet's name
'
'
' RETURN:  n > Index position
'         -1 > Not Found
'
Public Function SheetIndexPosition(ByVal pWorkbook As Workbook, _
                                   ByVal pSheetName As String) As Long

    Dim numberSheets As Long: numberSheets = pWorkbook.Worksheets.Count
    Dim i As Long

    SheetIndexPosition = -1

    pSheetName = Normalize(pSheetName)

    For i = 1 To numberSheets
        If Normalize(pWorkbook.Worksheets(i).Name) = pSheetName Then
            SheetIndexPosition = i
            Exit For
        End If
    Next i
End Function


'
' Activate the sheet, remove the filter mode and show the hidden columns.
'
' pSheet: Sheet
'
Public Sub SheetInitialize(ByVal pSheet As Worksheet)

    Dim lastColumn As Long: lastColumn = pSheet.UsedRange.Columns.Count

    pSheet.Activate

    Columns(NumberToLetter(lastColumn) & ":" & NumberToLetter(lastColumn)).Select

    Range(Selection, Selection.End(xlToLeft)).Select

    Selection.EntireColumn.Hidden = False

    If pSheet.FilterMode = True Then pSheet.ShowAllData

    pSheet.Cells(1, "A").Select
End Sub


'
' Show error message.
'
' pMessage: Error message.
' pErr (optional): ErrObject from error exception handler.
'
Public Sub ShowError(ByVal pMessage As String, _
                     Optional ByVal pErr As ErrObject = Nothing)

    If Not pErr Is Nothing Then
        pMessage = pMessage & vbCrLf & vbCrLf & pErr.Description
    End If

    MsgBox pMessage, vbCritical
End Sub


'
' Show information message.
'
' pLineFirst: Information message.
'
Public Sub ShowInfo(ByVal pMessage As String)
    MsgBox pMessage, vbInformation
End Sub


'
' Show warning message.
'
' pLineFirst: Warning message.
'
Public Sub ShowWarning(ByVal pMessage As String)
    MsgBox pMessage, vbExclamation
End Sub


'
' Raise exception with a custom mesasge.
'
Public Sub Throw(Optional ByVal pMessage As String = "")
    If pMessage <> "" Then Err.Description = pMessage
    Err.Raise 1
End Sub


'
' Join multiple values in a collection, where each one can be of different type.
'
' No arguments      = []
' "a", 1            = ["a", 1]
' "a", ["b", "c"]   = ["a", "b", "c"]
' ["b", "c"]        = ["b", "c"]
'
Function ToCollection(ParamArray pArray() As Variant) As Collection

    Dim i As Variant, j As Variant

    Set ToCollection = New Collection

    For Each i In pArray
        If TypeOf i Is Collection Or IsArray(i) Then
            For Each j In i
                ToCollection.add j
            Next j
        Else
            ToCollection.add i
        End If
    Next i
End Function


'
' Trim a string and convert it to lower case.
'
' pText: Text
'
' RETURN: Processed text
'
Public Function TrimLower(ByVal pText As String) As String
    TrimLower = LCase(Trim(pText))
End Function


'
' Trim a string and convert it to upper case.
'
' pText: Text
'
' RETURN: Processed text
'
Public Function TrimUpper(ByVal pText As String) As String
    TrimUpper = UCase(Trim(pText))
End Function


'
' Unzip a file.
'
' pFile: Input file to unzip
' pPathOutput (optional): Output folder
'
' RETURN: True  > OK
'         False > NOK
'
Public Function Unzip(ByVal pFile As String, _
                      Optional ByVal pPathOutput As String = "") As Boolean

    Dim file As Variant: file = pFile
    Dim pathOutput As Variant

    If pPathOutput = "" Then
        pathOutput = Left(pFile, Len(pFile) - Len(FileExtension(pFile)) - 1) & Application.PathSeparator
        If Dir(pathOutput, vbDirectory) = "" Then MkDir (pathOutput)
    Else
        If Right(pathOutput, 1) <> Application.PathSeparator Then pathOutput = pathOutput & Application.PathSeparator
    End If

    On Error GoTo ErrorHandler

    Dim app As Object: Set app = CreateObject("Shell.Application")
    app.Namespace(pathOutput).CopyHere app.Namespace(file).items

    Unzip = True
ErrorHandler:
End Function
