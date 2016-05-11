Function FormatSQL(Q As String, Optional CamposXLinea As Integer = 5) As String
    'Formatea una consulta en SQL, agregando identación y Saltos de linea.
    'Q: Query, o consulta a formatear
    'CamposXLinea: Numero de campos por linea que llevará el Select
    'Retorna: Una consulta formateada
    
    Dim sCampos As String
    Dim Contador As Integer
    Dim aCampos() As String
    Dim miCampo As Variant
    
    Dim sSelect As String
    Dim sFrom As String
    Dim sOrder As String
    
    sCampos = Between(Q, "SELECT", "FROM")
    sCampos = Replace(sCampos, " AS ", "**AS**")
    sCampos = Replace(sCampos, " ", "")
    sCampos = Replace(sCampos, "**AS**", " AS ")
    aCampos = Split(sCampos, ",")
    Contador = 0
    
    'SELECT
    For Each miCampo In aCampos
        Contador = Contador + 1
        sSelect = sSelect & miCampo & ", "
        If (Contador Mod CamposXLinea) = 0 Then
            sSelect = sSelect & vbNewLine & vbTab
        End If
    Next
    sSelect = Mid(sSelect, 1, Len(sSelect) - 2)
    sSelect = sSelect & " "
    sSelect = "SELECT " & sSelect & vbNewLine
    
    'FROM
    If InStr(1, Q, " WHERE ") Then
        sFrom = Between(Q, " FROM ", " WHERE ")
    ElseIf InStr(1, Q, " ORDER ") Then
        sFrom = Between(Q, " FROM ", " ORDER ")
    End If
    sFrom = "FROM " & sFrom & " "
    sFrom = Replace(sFrom, " LEFT ", " " & vbNewLine & vbTab & "LEFT ")
    sFrom = Replace(sFrom, " INNER ", " " & vbNewLine & vbTab & "INNER ")
    sFrom = Replace(sFrom, " RIGHT ", " " & vbNewLine & vbTab & "RIGHT ")
    sFrom = sFrom & vbNewLine
    
    'ORDER BY
    If InStr(1, Q, " ORDER ") Then
        sOrder = Mid(Q, InStr(1, Q, " ORDER ") + 1) & vbNewLine
    End If
    
    FormatSQL = sSelect & sFrom & sOrder
End Function


Function FormatSQL_Test()
    Dim Q As String
    Q = "Select IDObjetoMeta, Objeto, IDObjeto, IDObjetoDet, MetaNombre, MetaValor, tbaObjeto2.Descripcion FROM tbaObjetoMeta LEFT OUTER JOIN tbaObjeto2 ON tbaObjeto2.ID = tbaObjetoMeta.IDObjeto WHERE IDObjeto=10 ORDER BY Objeto"
    Debug.Print FormatSQL(Q, 4)
    
    'Imprime:
    'SELECT IDObjetoMeta, Objeto, IDObjeto, IDObjetoDet, 
    '  MetaNombre, MetaValor, tbaObjeto2.Descripcion 
    'FROM tbaObjetoMeta 
    '  LEFT OUTER JOIN tbaObjeto2 ON tbaObjeto2.ID = tbaObjetoMeta.IDObjeto 
    'ORDER BY Objeto
End Function
