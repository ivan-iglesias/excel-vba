# Convención de Nombres

El código de este repositorio sigue la siguiente convención de nombres. No es necesario su uso pero si recomendable para mantener la consistencia con el código de los ficheros helpers.

## Variables y constantes

Partimos que un fichero excel puede contener diferentes macros. Cada una de ellas ira en un módulo independiente sin poder acceder a otros, por lo que no usaremos las variables globales. Si queremos usar una variable global en el mismo módulo, para evitar conflictos con otros módulos las declararemos como privadas (Const es privada por defecto).

Como se indica a continuación, solo escribiremos en mayusculas las constantes. Usando la plantilla de macros existen dos excepciones, LOG y CONFIG. Son variables de módulo pero escritas como si fuesen constantes.

```vb
' Constant (snake_case with uppercase)
Const MY_CONSTANT as string = "value"

' Module variable (CamelCase)
Private BookName As String

' Function variable (lowerCamelCase)
Dim bookName As String

```

Al declarar **colecciones**, si son usadas como listado (puede tener elementos duplicados) el nombre tendrá el prefijo 'lst', en el caso de ser usada como diccionario (sin duplicados), tendrá 'dct'.

```vb
' List
Dim lstNames As Collection

' Dicctionary
Dim dctNames As Collection
```

## Funciones

El nombre de las **funciones** estará escrito en UpperCamelCase. Los argumentos que tengan irán también en UpperCamelCase pero su nombre con el prefijo 'p', indicando que la variable es un parametro.

```vb
Private Function GetPopulation(ByVal pFile As String, _
                               ByRef pDctData As Collection) As Integer
    ...
End Function
```

## Operadores & y +

Siempre usaremos el operador & cuando concatemnemos cadenas de texto, reservando el + para hacer sumas. Usar el operador + puede dar lugar a problemas si no se usa adecuadamente.

```vb
var1 = "10.01"
var2 = 11

var1 + var2   'result = 21.01
var1 & var2   'result = 10.0111
```
