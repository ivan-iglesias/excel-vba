Option Explicit

'
' Store the configuration sheet data into a collection. Due to it's simplicity,
' we do not perform any kind of validation or error handler
'
' pFileName: File name where is the configuration sheet
' pColumnKey: Key column position
' pColumnValue: Value column position
' pSheetName (optional): Configuration sheet name
'
' RETURN: Collection
'
Public Function ReadConfiguration(ByVal pFileName As String, _
                                  ByVal pColumnKey As Variant, _
                                  ByVal pColumnValue As Variant, _
                                  Optional ByVal pSheetName As String = "config") As Collection

    Dim sheet As Worksheet
    Dim line As Long
    Dim key As String
    Dim value As Variant

    Set ReadConfiguration = New Collection

    Set sheet = Workbooks(pFileName).Sheets(pSheetName)

    line = 2
    Do While Not IsEmpty(sheet.Cells(line, pColumnKey))
        key = TrimUpper(sheet.Cells(line, pColumnKey).value)
        value = sheet.Cells(line, pColumnValue).value

        ' If input value is a text, trim spaces
        If IsNumeric(value) Then value = Trim(value)

        ' Format path variables
        If key = "PATH_PROJECT" Then
            If Right(value, 1) <> "\" Then value = value & "\"

        ElseIf Left(key, 4) = "PATH" Then
            If Left(value, 1) = "\" Then value = Right(value, Len(value) - 1)
            value = ReadConfiguration.item("PATH_PROJECT") & value
        End If

        ReadConfiguration.add value, key

        line = line + 1
    Loop
End Function


'
' Check if the given files in the configuration sheet exists. If not, It asks
' the current location of each one.
'
' RETURN: True  > All files exists
'         False > NOK
'
Public Function FilesExists(ByRef pConfig As Collection, _
                            pKeys As Variant) As Boolean
    Dim key As Variant
    Dim file As String

    For Each key In pKeys
        file = pConfig(key)

        If Not FileExists(file) Then
            file = SelectFile("Select the file '" & FileName(file) & "' with key '" & key & "'")
            If file = "" Then Exit Function
            pConfig.Remove (key)
            pConfig.add file, key
        End If
    Next

    FilesExists = True
End Function
