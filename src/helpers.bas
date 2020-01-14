Option Explicit

'
' Add value with the given key if the does not exists in the collection.
'
' pCollection: Collection in which to add the item
' pKey: Value key
' pValue (optional): If the item is nothing, add the key as value
'
' RETURN: True  > OK
'         False > NOK
'
Public Function AddIfNotExists(ByRef pCollection As Collection, _
                               ByVal pKey As String, _
                               Optional ByVal pValue As Variant) As Boolean

    If IsMissing(pValue) Then pValue = pKey

    If Not ExistsKey(pCollection, pKey) Then
        pCollection.Add pValue, pKey
        AddIfNotExists = True
    End If
End Function

'
' Create an array of a given length.
'
' pLength: Array length
' pValue: Value to initialize the array
'
' RETURN: Array
'
Public Function ArrayInitialize(ByVal pLength As Long, _
                                ByVal pValue As Variant) As Variant()

    Dim myArray() As Variant
    Dim i As Long

    ReDim myArray(0 To pLength) As Variant

    For i = 0 To pLength
        myArray(i) = pValue
    Next

    ArrayInitialize = myArray
End Function

Public Function ArrayLen(pArray As Variant) As Long
    ArrayLen = UBound(pArray) - LBound(pArray) + 1
End Function

'
' Join array values.
'
' pArray: Array
' pDelimiter (optional): Delimiter between values
'
' RETURN: String
'
Public Function ArrayToString(ByRef pArray() As Variant, _
                              Optional ByVal pDelimiter As String = ";") As String

    ArrayToString = Join(pArray, pDelimiter)
End Function

'
' Join values from a collection.
'
' pCollection: Collection of strings or numbers
' pDelimiter (optional): Delimiter between values
'
' RETURN: String
'
Public Function CollectionToString(ByVal pCollection As Collection, _
                                   Optional ByVal pDelimiter As String = ";") As String

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
        collectionA.Add collectionB.item(i)
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
' Evaluate mathematical expression.
'
' RETURN: variant
'
Public Function Eval(ByVal pExpresion As String) As Variant
    On Error GoTo ErrorHandler
    pExpresion = Replace(pExpresion, ",", ".")
    Eval = Evaluate(pExpresion)
    If IsError(Eval) Then Eval = "error"
    Exit Function
ErrorHandler:
End Function

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
' Close excel file by name.
'
' pName: File's name
' pNumberBooks: Number of files currently open.
' pSave: True: Save changes
'        False: Not save changes
'
' RETURN: True  > OK
'         False > NOK
'
Public Function ExcelCloseByName(ByVal pName As String, _
                                 ByRef pNumberBooks As Long, _
                                 Optional ByVal pSave As Boolean = False) As Boolean

    On Error GoTo errHandler

    Dim i As Long

    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = pName Then
            Workbooks(i).Close savechanges:=pSave
            pNumberBooks = pNumberBooks - 1
            Exit For
        End If
    Next

    ExcelCloseByName = True
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
        Application.Workbooks.Add
        Workbooks(idxExcel).SaveAs fileName:=pFile
        ExcelCreate = True
    End If

    Exit Function
errHandler:
End Function

'
' Open the given text file.
'
' pFile: File to open
' pBooksOpen (optional): Number of open files, it is increased if the process ends successfully
' pFileNew (optional): File copy as txt (when opening csv files)
' pTab (optional): True = Use tab delimiter
' pSemicolon (optional): True = Use semicolon delimiter
' pComa (optional): True = Use comma delimiter
' pSpace (optional): True = Use space delimiter
'
' RETURN: True  > OK
'         False > NOK
'
Public Function ExcelOpen(ByVal pFile As String, _
                          Optional ByRef pBooksOpen As Long = 0, _
                          Optional ByRef pFileNew As String = "", _
                          Optional ByVal pTab As Boolean = False, _
                          Optional ByVal pSemicolon As Boolean = False, _
                          Optional ByVal pComa As Boolean = False, _
                          Optional ByVal pSpace As Boolean = False) As Boolean

    Const ARRAY_LENGTH As Long = 300

    On Error GoTo errHandler

    Dim extension As String: extension = LCase(GetFileExtension(pFile))

    ' Excel file
    If extension = "xls" Or extension = "xlsx" Then
        Workbooks.Open fileName:=pFile, UpdateLinks:=xlUpdateLinksAlways ', corruptload:=xlRepairFile
        pBooksOpen = pBooksOpen + 1
        ExcelOpen = True
        Exit Function
    End If

    ' CSV/TXT file
    If extension = "csv" Then
        pFileNew = Left(pFile, Len(pFile) - 3) & "txt"
        If ExistsFile(pFileNew) Then FileDelete (pFileNew)
        Call FileCopy(pFile, pFileNew)
        pFile = pFileNew
    ElseIf extension <> "txt" Then
        Exit Function
    End If

    ' An array containing parse information for individual columns of data.
    ' The first element is the column number, and the second element is then constants
    ' (XlColumnDataType) specifying how the column is parsed. We use 2 (xlTextFormat)
    ' to avoid data loss when parsing codes starting with 0.
    Dim myArray(0 To ARRAY_LENGTH) As Variant
    Dim i As Long
    For i = 0 To ARRAY_LENGTH
        myArray(i) = Array(i + 1, 2)
    Next

    ' Open the file
    Workbooks.OpenText fileName:=pFile, _
                       StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                       ConsecutiveDelimiter:=False, Tab:=pTab, Semicolon:=pSemicolon, Comma:=pComa, Space:=pSpace, Other:=False, _
                       fieldInfo:=myArray, Local:=True

    pBooksOpen = pBooksOpen + 1
    ExcelOpen = True
    Exit Function
errHandler:
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
' Check if exists a file.
'
' pFile: File's full path to check
'
' RETURN: True  > Exists
'         False > Not exists
'
Public Function ExistsFile(ByVal pFile As String) As Boolean
    If Dir(pFile) <> "" Then ExistsFile = True
End Function

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
' Create a copy of a file with unix format.
'
' pFileInput: File's full path to translate
'
' RETURN: OK  > New file path full path
'         NOK > Empty
'
Public Function FileToUnix(ByVal pFileInput As String) As String
    Dim text As String

    FileToUnix = Replace(pFileInput, ".txt", "_unix.txt")

On Error GoTo errHandler

    Open pFileInput For Input As #1
    Open FileToUnix For Output As #2

    Do Until EOF(1)
        Line Input #1, text
        text = Replace(text, vbCrLf, Chr(10))
        Print #2, Chr(10); text;
    Loop

    Close #2
    Close #1
    Exit Function
errHandler:
    FileToUnix = ""
End Function

'
' Create a new excel file.
'
' pPath: Folder's parent path
' pName: Folder's name
' pWithDateTime (optional): Add timestamp to folder's name
'
' RETURN: String, OK  > Folder full path
'                 NOK > ""
'
Public Function FolderCreate(ByVal pPath As String, _
                             ByVal pName As String, _
                             Optional ByVal pWithDateTime As Boolean = True) As String

    On Error GoTo errHandler

    pPath = FolderEndingDelimiter(pPath)

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
        pPath = FolderEndingDelimiter(pPath)
        Kill pPath & "*.*"
        RmDir pPath
        FolderDelete = True
    On Error GoTo 0
End Function

'
' Check if directory path ends with 'PathSeparator'.
'
' pPath: path to validate
'
' RETURN: Given directory. If it has not the ending delimiter, we add it
'
Public Function FolderEndingDelimiter(ByVal pPath As String) As String

    If Right(pPath, 1) <> Application.PathSeparator Then pPath = pPath & Application.PathSeparator

    FolderEndingDelimiter = pPath
End Function

'
' Get current date time.
'
' pFormatDate (optional): Date format
' pFormatTime (optional): Time format
'
' RETURN: Date time with given format
'
Public Function GetDateTime(Optional ByVal pFormatDate As String = "yyyMMdd", _
                            Optional ByVal pFormatTime As String = "HHmm") As String

    GetDateTime = Format(Date, pFormatDate) & "_" & Format(Time, pFormatTime)
End Function

'
' Get the extension from file's full path.
'
' pFile: File full path
'
' RETURN: String
'
Public Function GetFileExtension(ByVal pFile As String) As String
    GetFileExtension = Right(pFile, Len(pFile) - InStrRev(pFile, "."))
End Function

'
' Get the file name from its full path.
'
' pFile: File full path
'
' RETURN: String
'
Public Function GetFileName(ByVal pFile As String) As String
    GetFileName = Right(pFile, Len(pFile) - InStrRev(pFile, Application.PathSeparator))
End Function

'
' Get the file path from its full path.
'
' pFile: File full path
'
' RETURN: String
'
Public Function GetFilePath(ByVal pFile As String) As String
    GetFilePath = Left(pFile, InStrRev(pFile, Application.PathSeparator))
End Function

'
' Get the files of a given folder.
'
' pPath: Folder's path
'
' RETURN: Collection of files
'
Public Function GetFolderFiles(ByVal pPath As String) As Collection
    Set GetFolderFiles = New Collection

    On Error GoTo errHandler

     pPath = FolderEndingDelimiter(pPath)

    Dim file As String: file = Dir(pPath)

    Do While file <> ""
        GetFolderFiles.Add pPath & file
        file = Dir()
    Loop
errHandler:
End Function

'
' Get the week number from a given date.
'
' RETURN: Week number
'
Public Function GetWeek(ByVal pDate As String, _
                        Optional ByVal pPrefix As String = "") As String

    If Len(pDate) = 8 Then
        pDate = Left(pDate, 4) & "/" & Mid(pDate, 5, 2) & "/" & Right(pDate, 2)
    End If

    GetWeek = pPrefix & Format(CDate(pDate), "ww", vbMonday, vbFirstFourDays)
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
    Dim fileName As String
    Dim wb As Variant

    On Error Resume Next

    fileName = LCase(GetFileName(pFile))

    For Each wb In Workbooks
        If LCase(wb.Name) = fileName Then
            ExcelIsOpen = True
            Exit Function
        End If
    Next wb
End Function

'
' Trim a string and convert it to lower case.
'
' pText: Text
' pReplaceSpaceWithUnderscore (optional): Replace spaces with underscores
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
' Concatenate a range of cells.
'
' pRange: Range of cells to concatenate
' pDelimiter (optional): Delimiter between values
'
' RETURN: String
'
Public Function RangeToString(ByVal pRange As Range, _
                              Optional ByVal pDelimiter As String = ",") As String

    Dim item As Variant

    For Each item In pRange
        RangeToString = RangeToString & item.Value & pDelimiter
    Next

    If RangeToString <> "" Then
        RangeToString = Left(RangeToString, Len(RangeToString) - Len(pDelimiter))
    End If
End Function

'
' Round number.
'
' pNumber: Number to be rounded
' pDecimals: Number of decimals
'
' RETURN: Number
'
Public Function RoundUp(ByVal pNumber As Double, _
                        Optional ByVal pDecimals As Long = 0) As Double

    RoundUp = WorksheetFunction.RoundUp(pNumber, pDecimals)
End Function

'
' Returns the column index position of a field.
'
' pSheet: Sheet where is the field to search
' pName: Text to search
' pRow (optional): Row where the text must be
' pMessage (optional): If different to 'na', store the fields name if missing
'
' RETURN:  n > Column index position
'         -1 > Not Found
'
Public Function SearchColumn(ByVal pSheet As Worksheet, _
                             ByVal pName As String, _
                             Optional ByVal pRow As Long = 1, _
                             Optional ByRef pMessage As String = "na", _
                             Optional ByVal pContains As Boolean = False) As Long

    SearchColumn = -1

    Dim pNameNormalize As String: pNameNormalize = Normalize(pName)

    If pNameNormalize = "" Then Exit Function

    Dim currentColumnName As String
    Dim pNumberColumns As Long
    Dim i As Long

    pNumberColumns = pSheet.UsedRange.Rows(pRow).Columns.Count - 1

    i = 1
    Do While Not IsEmpty(pSheet.Cells(pRow, i)) Or i <= pNumberColumns

        currentColumnName = Normalize(pSheet.Cells(pRow, i).Value)

        If currentColumnName = pNameNormalize Or _
           (pContains And InStr(currentColumnName, pNameNormalize) <> 0) Then
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
        If Normalize(pSheet.Cells(i, pColumn).Value) = pName Then
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

    Dim idx As Long: idx = SheetIndexPosition(pWorkbook, pSheetName)

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
' pFontSize (optional): Font's size
' pColumnWidth (optional): Columns width
' pRowHeight (optional): Rows height
'
Public Sub SheetBasicStyle(ByVal pSheet As Worksheet, _
                           Optional pFontSize As Long = 11, _
                           Optional pColumnWidth As Long = -1, _
                           Optional pRowHeight As Long = -1)

    With pSheet
        .Cells.Font.Size = pFontSize
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
' pMessage (optional): Exception message
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
                ToCollection.Add j
            Next j
        Else
            ToCollection.Add i
        End If
    Next i
End Function

'
' Create a dictionary from a given collection.
'
' pCollection: Collection (List/Dictionary)
'
' RETURN: Collection (Dictionary)
'
Function ToDictionary(ByVal pCollection As Collection) As Collection
    Dim i As Variant

    Set ToDictionary = New Collection

    For Each i In pCollection
        If Not ExistsKey(ToDictionary, i) Then
            ToDictionary.Add i, i
        End If
    Next
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
' pPath (optional): Output folder
'
' RETURN: True  > OK
'         False > NOK
'
Public Function Unzip(ByVal pFile As String, _
                      Optional ByVal pPath As String = "") As Boolean

    Dim file As Variant: file = pFile

    If pPath = "" Then
        pPath = Left(pFile, Len(pFile) - Len(GetFileExtension(pFile)) - 1) & Application.PathSeparator
        If Dir(pPath, vbDirectory) = "" Then MkDir (pPath)
    Else
        pPath = FolderEndingDelimiter(pPath)
    End If

    On Error GoTo ErrorHandler

    Dim app As Object: Set app = CreateObject("Shell.Application")
    app.Namespace(pPath).CopyHere app.Namespace(file).items

    Unzip = True
ErrorHandler:
End Function

'
' Write datetime value into cell.
'
Public Sub WriteDate(ByVal pCell As Object, _
                     ByVal pValue As String)

    If IsDate(pValue) And Len(pValue) = 10 Then
        pCell = CDate(pValue)
    Else
        pCell = pValue
    End If
End Sub
