## Imágenes

### Guardar rango como imagen

En el caso de que queramos formar una imagen compuesta por múltiples gráficas o imagenes, incluso añadiendo textos a celdas como cabeceras, podemos crear una imagen del área deseada y pegarla como una nueva imagen en el excel o en un fichero externo (jpg, png).

```vb
Dim sheet As Worksheet
Set sheet = Workbooks("workbook_name").Sheets("sheet_name")

sheet.range("B2:O20").CopyPicture xlScreen, xlPicture

' Save the picture in the current sheet
sheet.Paste Destination:=sheet.range("B21")

' Save the picture to file
Dim objChart As Chart: Set objChart = Charts.Add
With objChart
	.Paste
	.Export Filename:="C:\picture.jpg", Filtername:="JPG"
End With
```
