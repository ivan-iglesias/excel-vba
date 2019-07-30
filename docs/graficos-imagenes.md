# Gráficos e Imagenes

## Crear

Podemos crear de forma sencilla un gráfico pasando el rango de celdas que contienen la información, y una vez creado, pegarlo como si fuese una imagen.

```vb
Dim sheet As Worksheet
Dim dataRange As Range
Dim objChart As ChartObject

Set sheet = Workbooks("workbook_name").Sheets("sheet_name")
Set dataRange = sheet.Range("A1:C20")

' Create a chart
Set objChart = sheet.ChartObjects.add(50, 40, 500, 300)

objChart.Chart.ChartWizard Source:=dataRange, _
	Gallery:=xlLineStacked, Format:=5, PlotBy:=xlColumns, _
	Title:="Example", HasLegend:=True

With objChart
    .chart.ChartTitle.Font.Size = 13
    .chart.ChartTitle.Font.Bold = False
    .Name = "example_chart"
    .Left = sheet.Range("D2").Left
    .Top = sheet.Range("D2").Top
End With

' Export chart as picture
objChart.Activate
ActiveChart.ChartArea.Copy

sheet.Activate
sheet.Cells(2, "M").Select
sheet.Pictures.Paste
sheet.Pictures(sheet.Pictures.Count).Name = "example_chart_picture"
```

## Borrar

Podemos borrar las gráficas o imagenes de un excel de una en una o todas a la vez si no nos interesa hacer algún tipo de filtrado.

```vb
dim sheet as Worksheet
dim item as variant

Set sheet = Workbooks("workbook_name").Sheets("sheet_name")

' Delete all charts one by one
For Each item In sheet.ChartObjects
	item.Delete
Next

' Delete all charts
sheet.ChartObjects.Delete

' Delete all pictures one by one
For Each item In sheet.Pictures
	item.Delete
Next

' Delete all pictures
sheet.Pictures.Delete
```

## Guardar área del excel como imagen

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

# Tipos de gráficos

A continuación listo alguno de los gráficos disponibles. Para mayor información acuda a la página oficial de [Microsoft](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlcharttype?view=excel-pia) en la que encontrará el listado completo de opciones disponibles que ofrece ```XlChartType```.

| Name      					| Description 				 	|
| ----------------------------- | -----------------------------	|
|	xlArea						|	Area					 	|
|	xlAreaStacked				|	Stacked Area			 	|
|	xlAreaStacked100			|	100% Stacked Area	 	 	|
|	xlBarClustered				|	Clustered Bar			 	|
|	xlBarStacked				|	Stacked Bar				 	|
|	xlBarStacked100				|	100% Stacked Bar		 	|
|	xlColumnClustered			|	Clustered Column		 	|
|	xlColumnStacked				|	Stacked Column			 	|
|	xlDoughnut					|	Doughnut				 	|
|	xlLine						|	Line					 	|
|	xlLineMarkers				|	Line with Markers			|
|	xlPie						|	Pie							|
|	xlRadar						|   Radar				    	|
|	xlStockOHLC					|   Open-High-Low-Close			|
|	xlXYScatter					|   Scatter						|
|	xlXYScatterLines			|   Scatter with Lines		 	|
|	xlXYScatterSmooth	        |   Scatter with Smoothed Lines	|


