
# VBA Excel

Conjunto de funciones diseñadas para facilitar el desarrollo de macros Excel de forma eficiente y ordenada.

A continuación listo los ficheros existentes, los cuales habrá que importar al proyecto en función de las necesidades.

- ```helpers.bas```, funciones para el desarrollo de macros.
- ```helpers_word.bas```, permite adjuntar imágenes, tablas y textos en documentos word.
- ```main_template.bas```, plantilla inicial de nueva macro.
- ```helpers_template.bas```, funciones complementarias cuando se hace uso de `main_template`.

> Los cambios realizados en cada versión están documentados en las [notas de vesión](https://github.com/ivan-iglesias/excel-vba/releases).

## Ejemplos de código

- [Cheatsheet](https://github.com/ivan-iglesias/excel-vba/blob/master/docs/cheatsheet.md)
- [Creary eliminar gráficos](https://github.com/ivan-iglesias/excel-vba/blob/master/docs/chart.md)
- [Guardar rango como imágen](https://github.com/ivan-iglesias/excel-vba/blob/master/docs/picture.md)
- [Exportar libro o rango a PDF](https://github.com/ivan-iglesias/excel-vba/blob/master/docs/pdf.md)

## Convención de Nombres

El código de este repositorio sigue la siguiente convención de nombres. No es necesario su uso pero si recomendable para mantener la consistencia con el código de los ficheros `helpers`.

### Variables y constantes

Partimos que un fichero excel puede contener diferentes macros. Cada una de ellas ira en un módulo independiente sin poder acceder a otros, por lo que no usaremos las variables globales. Si queremos usar una variable global a nivel de módulo, para evitar conflictos con otros módulos las declararemos como privadas (`const` es privada por defecto).

Como se indica a continuación, solo escribiremos en mayúsculas las constantes. Usando la plantilla de macros existen dos excepciones, LOG y CONFIG, las cuales son variables de módulo pero escritas como si fuesen constantes.

```vb
' Constant (snake_case with uppercase)
Const MY_CONSTANT as string = "value"

' Module variable (CamelCase)
Private BookName As String

' Function variable (lowerCamelCase)
Dim bookName As String
```

Al declarar **colecciones**, si son usadas como listado (puede tener elementos duplicados) el nombre tendrá el prefijo **lst**, en el caso de ser usada como diccionario (sin duplicados), tendrá **dct**.

```vb
' List
Dim lstNames As Collection

' Dicctionary
Dim dctNames As Collection
```

Al declarar una variable como `variant`, para recorrer una colección, podremos añadir opcionalmente el prefijo `v`.

```vb
Dim item As Collection
Dim vItem As Collection

for each item in lstNames
    ' ...
next

for each vItem in lstNames
    ' ...
next
```

### Funciones

El nombre de las **funciones** estará escrito en *UpperCamelCase*. Los argumentos que tengan irán también en *UpperCamelCase* pero su nombre con el prefijo 'p', indicando que la variable es un parámetro.

```vb
Private Function GetPopulation(ByVal pFile As String, _
                               ByRef pDctData As Collection) As Integer
    ...
End Function
```

### Operadores & y +

Siempre usaremos el operador `&` cuando concatenemos cadenas de texto, reservando el `+` para hacer sumas. Usar el operador `+` puede dar lugar a problemas si no se usa adecuadamente.

```vb
var1 = "10.01"
var2 = 11

var1 + var2   'result = 21.01
var1 & var2   'result = 10.0111
```
