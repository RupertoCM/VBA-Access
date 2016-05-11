Function Between(sCadena As String, Coincidencia1 As String, Coincidencia2 As String) As String
    'Devuelve una cadena que se encuentra entre dos coincidencias
    'En otros lenguajes es conocina como: StringBetween
    Dim PosC1, PosC2, nLetras As Integer
    
    PosC1 = InStr(1, sCadena, Coincidencia1)
    If PosC1 > 0 Then
        PosC1 = PosC1 + Len(Coincidencia1)
    Else
        GoTo NOCoincide
    End If
    
    PosC2 = InStr(PosC1, sCadena, Coincidencia2)
    If PosC2 <= 0 Then
        GoTo NOCoincide
    End If
    
    nLetras = PosC2 - PosC1
    
    Between = Mid(sCadena, PosC1, nLetras)
    
    Exit Function
NOCoincide:
    Between = ""
End Function

Function Between_Test()
  debug.print Between("Mi nombre es: [Ruperto Coronado]","[","]")
  'Imprime:
  'Ruperto Coronado
End Function
