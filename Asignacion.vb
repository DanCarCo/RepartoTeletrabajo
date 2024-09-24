ub AsignarNombresCompleto()
    Dim ws As Worksheet
    Dim celdasGrupo1 As Variant
    Dim nombres As Variant
    Dim nombresAdicionales As Variant
    Dim nombresAdicionales2 As Variant
    Dim nombresAdicionales3 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim nombre As String
    Dim rango As Range
    Dim celda As Range

    ' Establece la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Hoja 2")

    ' Asignar nombres desde L2, L3, L4 a los rangos del primer grupo
    nombres = Array(ws.Range("L2").Value, ws.Range("L3").Value, ws.Range("L4").Value)
    celdasGrupo1 = Array("C5", "E5", "G5", "C10", "E10", "G10", _
                         "C15", "E15", "G15", "C20", "E20", "G20", _
                         "D4", "F4", "D9", "G9", "D12", "G14", _
                         "D19", "G19", "D24", "G24", "C25", "E25", "G25")
    
    For i = LBound(celdasGrupo1) To UBound(celdasGrupo1)
        nombre = Join(nombres, ", ")
        ws.Range(celdasGrupo1(i)).Value = nombre
    Next i

    ' Asignar nombres adicionales desde M2 y M3 a los rangos del segundo grupo
    nombresAdicionales = Array(ws.Range("M2").Value, ws.Range("M3").Value)
    celdasGrupo1 = Array("C4:G4", "C9:G9", "C14:G14", "C19:G19", "C24:G24")
    
    For i = LBound(celdasGrupo1) To UBound(celdasGrupo1)
        Set rango = ws.Range(celdasGrupo1(i))
        For Each celda In rango
            If IsEmpty(celda.Value) Then
                celda.Value = Join(nombresAdicionales, ", ")
            Else
                celda.Value = celda.Value & ", " & Join(nombresAdicionales, ", ")
            End If
        Next celda
    Next i

        ' Asignar nombres adicionales desde N2, N3, N4 a las celdas del cuarto grupo
    nombresAdicionales2 = Array(ws.Range("N2").Value, ws.Range("N3").Value, ws.Range("N4").Value)
    celdasGrupo1 = Array("C4", "D5", "E4", "F5", "G4", _
                          "C10", "D9", "E10", "F9", "G10", _
                          "C14", "D15", "E14", "F15", "G14", _
                          "C20", "D19", "E20", "F19", "G20", _
                          "C24", "D25", "E24", "F25", "G24")
    
    For i = LBound(celdasGrupo1) To UBound(celdasGrupo1)
        Set celda = ws.Range(celdasGrupo1(i))
        If IsEmpty(celda.Value) Then
            celda.Value = Join(nombresAdicionales2, ", ")
        Else
            celda.Value = celda.Value & ", " & Join(nombresAdicionales2, ", ")
        End If
    Next i

    ' Asignar nombres adicionales desde O2, O3, O4 a las celdas del tercer grupo
    nombresAdicionales3 = Array(ws.Range("O2").Value, ws.Range("O3").Value, ws.Range("O4").Value)
    celdasGrupo1 = Array("C5", "D4", "E5", "F4", "G5", _
                          "C9", "D10", "E9", "F10", "G9", _
                          "C15", "D14", "E15", "F14", "G15", _
                          "C19", "D20", "E19", "F20", "G19", _
                          "C25", "D24", "E25", "F24", "G25")
    
    For i = LBound(celdasGrupo1) To UBound(celdasGrupo1)
        Set celda = ws.Range(celdasGrupo1(i))
        If IsEmpty(celda.Value) Then
            celda.Value = Join(nombresAdicionales3, ", ")
        Else
            celda.Value = celda.Value & ", " & Join(nombresAdicionales3, ", ")
        End If
    Next i
End Sub


