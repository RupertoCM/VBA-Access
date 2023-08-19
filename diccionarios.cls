'Funciones o utilidades con diccionarios para VBA de Access
'Para poder llamar al diccionario directamente; agrega la referencia al objeto: Microsoft Scripting Runtime en el menú Herramientas > Referencias

Public Sub Dict_Print_AsTable(d As Dictionary)
    'Imprime en la consola un diccionario en formato de tabla

    Dim item As Variant
    
    Dim sD As Dictionary 'subdict
    Dim dColumnas As New Dictionary 'Guardara: ColumName: MaxSize
    Dim clave As Variant
    Dim intTableWidth As Integer
    Dim subitem As Variant
    Dim sClave As String
    Dim sValor As String
    Dim sColumna As String
    
    'Obtener el nombre de la columna y tamaño máximo de los valores de columna. _
    Por ejemplo: Si el diccionario tiene 3 registros, se obtiene la longitud máxima del valor de la columna 'nombre'
    For Each item In d
        Set sD = d(item)
        
        For Each subitem In sD
            
            sColumna = subitem
            sValor = " " & sD(sColumna) & " "
            
            If dColumnas.Exists(sColumna) Then 'Si ya existe la columna, almacenar la mayor longitud
                If Len(sValor) > CInt(dColumnas(sColumna)) Then
                    intTableWidth = intTableWidth + Len(sValor) - dColumnas(sColumna)
                    dColumnas(sColumna) = Len(sValor)
                    
                End If
            Else 'Agregar al diccionario la columna, con la longitus mayor del nombre de la columna o del registro
                If Len(sColumna) > Len(sValor) Then 'Si len(columna) > len(valor)
                    dColumnas.Add sColumna, Len(sColumna) + 2
                Else
                    dColumnas.Add sColumna, Len(sValor)
                End If
                intTableWidth = intTableWidth + dColumnas(sColumna)
            End If
        Next
    Next
    
    intTableWidth = intTableWidth + dColumnas.count + 1
    'Imprimir las cabeceras de las columnas
    Dim sLineaSeparadora As String
    sLineaSeparadora = String$(intTableWidth, "-")
    
    '---------------------
    '|columna1|columna2...
    '---------------------
    Debug.Print sLineaSeparadora
    Debug.Print "|";
    For Each item In dColumnas
        Debug.Print TF(" " & item & " ", "L", dColumnas(item)) & "|";
    Next
    Debug.Print ""
    Debug.Print sLineaSeparadora
    
    
    Dim sAlineacion As String
    For Each item In d
        Set sD = d(item)
        
        Debug.Print "|"; 'El ; al final hace que no salte línea
        For Each subitem In sD
            sValor = sD(subitem)
            sAlineacion = "I"
            If IsNumeric(sValor) Then sAlineacion = "D" 'Si es número lo alineamos a la derecha
            sValor = " " & sValor & " "
            
            'Insertamos espacios a la derecha o izquierda. Con la fórmula: Cantidad Espacios=Longitud de la columna - longitud del valor
            If sAlineacion = "I" Then 'Izquierda
                Debug.Print sValor & Space(dColumnas(subitem) - Len(sValor)) & "|";
            Else 'Derecha
                Debug.Print Space(dColumnas(subitem) - Len(sValor)) & sValor & "|";
            End If
        Next
        Debug.Print ""
    Next
    Debug.Print sLineaSeparadora
End Sub


Sub Dict_Print_AsTable_Test()
    'Generamos un diccionario de Productos. El nombre de Cada elemento Producto será su ID. _
        Cada producto tiene id, nombre, costo y precio que se guardan en un subdiccionario, _
        Es buena práctica que cada clave tenga el nombre en minúscula pues los objetos diccionarios _
        son (por defecto) sensibles a mayúsculas y minúsculas 

    Dim dProductos As New Dictionary 'Diccionario principal (que agrupa productos)
    Dim dProducto As Dictionary 'Subdiccionario que gaurda cada producto
    
    Set dProducto = New Dictionary
    Dim i As Integer
    
    For i = 1 To 5
        Set dProducto = New Dictionary
        dProducto("id") = i
        dProducto("nombre") = "Producto " & i
        dProducto("costo") = Format(Int((150 - 10 + 1) * Rnd + 10), "fixed") 'Genera un costo aleatorio con formato fijo
        dProducto("precio") = Format(dProducto("costo") / 0.8, "fixed") 'Costo + 20% de utilidad
        dProductos.Add i, dProducto
    Next
    
    
    
    
    Dict_Print_AsTable dProductos

  'Imprime:
  '------------------------------------
  '| id| nombre     | costo  | precio |
  '------------------------------------
  '| 1 | Producto 1 |  68.00 |  85.00 |
  '| 2 | Producto 2 | 131.00 | 163.75 |
  '| 3 | Producto 3 | 121.00 | 151.25 |
  '| 4 | Producto 4 |  62.00 |  77.50 |
  '| 5 | Producto 5 | 145.00 | 181.25 |
  '------------------------------------
End Sub
