# Exportar a PDF

Para exportar un libro excel o rango a un fichero PDF, debemos hacer uso de función *ExportAsFixedFormat* desde el libro o rango deseado.

Exportar directamente un **rango**

```vb
Dim sheet As Worksheet
Dim dataRange As Range

Set sheet = Workbooks("workbook_name").Sheets("sheet_name")

Set dataRange = sheet.Range("A1:I30")

Call dataRange.ExportAsFixedFormat(Type:=xlTypePDF, FileName:="C:/output/file.pdf")
```

Exportar un **libro**

```vb
Dim wb as Workbook

Set wb = Workbooks("workbook_name")

Call wb.ExportAsFixedFormat Type:=xlTypePDF, _
    FileName:="C:/output/file.pdf", _
    Quality:=xlQualityStandard, _
    includedocproperties:=True, _
    IgnorePrintAreas:=False, _
    openafterpublish:=pOpenAfterPublish, _
    To:=2
```
