Option Explicit

' Module constants
Const LINE_HEADER = 1

' Module variables
Private LOG As Collection
Private CONFIG As Collection


' Main Function
Public Sub MainTemplate()
    Dim booksOpen As Long
    Dim file As String
    Dim dctData As New Collection

    Set LOG = New Collection
    Set CONFIG = ReadConfiguration(ActiveWorkbook.Name, "A", "B")

    Call EnableOptimization

    If Not CheckIfFilesExists(CONFIG, Array("PATH_DISTRICTS")) Then EndProcess


    ' Get input data from file
    file = SelectFile("Select input file", "xls")
    If file = "" Then EndProcess (booksOpen)
    If Not ExcelOpen(file, booksOpen) Then EndProcess
    If Not GetData(FileName(file), dctData) Then EndProcess (booksOpen)
    Call ExcelClose(booksOpen)


    Call ShowInfo("Execution completed")
    Call EnableOptimization(False)
End Sub

'
' Get data from the file.
'
Private Function GetData(ByVal pFileName As String, _
                         ByRef pDctData As Collection) As Boolean

    Dim sheet As Worksheet
    Dim line As Long
    Dim idxField As Long
    Dim temp As String

    On Error GoTo errHandler

    ' Select working sheet
    Set sheet = Workbooks(pFileName).Sheets(1)
    Call SheetInitialize(sheet)

    ' Search fields
    idxField = SearchColumn(sheet, CONFIG("FIELD_NAME"), pMessage:=temp)

    If temp <> "" Then Throw (pFileName & " - Could not find fields: " & vbCrLf & vbCrLf & temp)

    ' Process data
    line = LINE_HEADER
    Do While Not IsEmpty(sheet.Cells(line, idxField))
        temp = TrimUpper(sheet.Cells(line, idxField))
        If Not ExistsKey(pDctData, temp) Then pDctData.add temp, temp
        line = line + 1
    Loop

    GetData = True
    Exit Function
errHandler:
    Call ShowError("Error in GetData.", pErr:=Err)
End Function
