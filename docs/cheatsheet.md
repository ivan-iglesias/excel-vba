### Estilos

```vb
' Colores, fuente, etc.
With sheet
	.Cells.Font.Size = 10
	.Cells.Font.name = "Calibri"
	.Columns().HorizontalAlignment = xlLeft
	.Columns().VerticalAlignment = xlCenter
	.Rows(1).HorizontalAlignment = xlCenter
    .Rows(1).Orientation = 90
    .Range("A1:A10").Orientation = 90

    .Rows.AutoFit
    .Rows().RowHeight = 15
    .Rows(1).WrapText = True
    .Columns().AutoFit
    .Columns().ColumnWidth = 15
    .Columns("A").ColumnWidth = 20

	' vbBlack, vbBlue, vbCyan, vbGreen, vbMagenta, vbRed, vbWhite, vbYellow
	.Range("A1:A10").Font.Color = vbRed
    .Range("A1:A10").Interior.Color = RGB(211, 211, 211)
    .Range("A1:A10").Font.Bold = True
End With

' Formatos de celda
sheet.Cells(line, "A").NumberFormat = "dd/mm/yyyy"
sheet.Cells(line, "A").NumberFormat = "0.00"
sheet.Cells(line, "A").NumberFormat = "@"

' Inmovilizar paneles
sheet.Activate
With ActiveWindow
	.SplitColumn = 0
	.SplitRow = 1
End With
ActiveWindow.FreezePanes = True
```

### Añadir y Copiar Pestañas

```vb
'Añadir sheet
wb.Worksheets.Add(Before:=Worksheets(1)).Name = "new"
wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "new"

'Copiar y cambiar nombre a sheet
sheetIn.Copy Before:=wb.Sheets(1)
ActiveSheet.Name = "new"
```

### Copiar Líneas

```vb
' Copiar linea por valor
sheetIn.Range("A1:C5").Copy
sheetOut.Range("A1").PasteSpecial xlPasteValues

' Copiar linea
sheet.Rows(1).EntireRow.Copy Destination:=wb.Sheets("out").Range("A1")

' Remplazar
sheet.Columns("A").Replace What:=",", Replacement:=";", SearchOrder:=xlByColumns, MatchCase:=True

' Copiar rango
sheetIn.Range("A2:B10").Copy Destination:= sheetOut.Range("A1")
```

### Eliminar Contenido

| DESCRIPCIÓN                           | COMANDO                     |
| ------------------------------------- | --------------------------- |
| Eliminar **contenido** y **formatos** | `sheet.Cells.Clear`         |
| Eliminar **contenido**, no formatos   | `sheet.Cells.ClearContents` |
| Eliminar solo **formatos**            | `sheet.Cells.ClearFormats`  |
| Eliminar **comentarios**              | `sheet.Cells.ClearComments` |
| Eliminar **comentarios**              | `sheet.Cells.ClearNotes`    |
| Eliminar **comentarios**              | `sheet.Cells.ClearOutline`  |

### Filtrado

```vb
' Activar filtrado
sheet.Range("A1").AutoFilter

' Filtrar y copiar resultado
sheet.Range("A:A").AutoFilter Field:=idx, Criteria1:=value, Operator:=xlFilterValues
sheet.AutoFilter.Range.Copy Destination:=Workbooks("out").sheets(1).Range("A1")

' Contar líneas filtradas
Sheet.AutoFilter.Range.Columns(idxCluster).SpecialCells(xlCellTypeVisible).Cells.count
```

### Ordenar Rango

```vb
' Contar número de líneas y columnas usadas
Dim lastLine as long
Dim lastLine as string

lastLine = sheet.Cells(sheet.Rows.Count, "A").End(xlUp).Row
lastColumn = sheet.Cells(1, sheet.Columns.Count).End(xlToLeft).Column

lastLine = sheet.UsedRange.Rows(sheet.UsedRange.Rows.count).Row
lastColumn = NumberToLetter(sheet.UsedRange.columns.column)

' Ordenar'
sheet.Range("A1:C10").Sort Key1:=Range("A1:A10"), Order1:=xlAscending, Header:=xlYes

sheet.Range("A1").Sort header:=xlYes, Key1:=sheet.Range(Cells(1, "B").Address), Order1:=xlDescending
```

### Formulas

```vb
' El separado en la formula depende del idioma del ordenador (español ",", ingles ";")
sheet.Cells(line, "A").FormulaLocal = "=SI(M" & line & "="""";"""";DIA.LAB(M" & line & ";0))"
sheet.Cells(line, "A").Formula = "=IF(2=2;5;6)"
sheet.Cells(line, "A").FormulaLocal = "=B" & line & "/24*3"
```

### Leer Fichero TXT/CSV

```vb
Dim file as string
Dim lineText as string
Dim lstLines as new Collection

Open file For Input As #1

Do Until EOF(1)
	Line Input #1, lineText
	lstLines.add Trim(lineText)
Loop

Close #1
```

### Añadir Comentario

```vb
Dim cmt As Comment
Dim header As String
dim text As String

header1 = "Tareas a realizar:" & vbLf
text1 = "- Levantarse" & vbLf & vbLf "- Hacer la cama"
header1 = "Tareas realizadas:" & vbLf
text2 = "- Compras"

Set cmt = sheet.cells(1, "A").AddComment(header1 & text1 & header2 & text2)

With cmt.Shape.TextFrame
    .Characters.Font.Bold = False
    .Characters(1, Len(header1)).Font.Bold = True
    .Characters(Len(header1) + Len(text1) + 1, Len(header2)).Font.Bold = True
    .AutoSize = True
End With
```

#### Varios

```vb
' Quitar compartido
If wb.MultiUserEditing Then wb.ExclusiveAccess
```

```vb
'YES/NO msgbox
If MsgBox("¿Quieres completar la plantilla?", vbYesNo + vbQuestion, "Plantilla") = vbYes Then

else

end if
```
