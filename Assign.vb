Sub AssignFullNames()
    Dim ws As Worksheet
    Dim group1Cells As Variant
    Dim names As Variant
    Dim additionalNames As Variant
    Dim additionalNames2 As Variant
    Dim additionalNames3 As Variant
    Dim i As Integer
    Dim j As Integer
    Dim name As String
    Dim range As range
    Dim cell As range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Assign names from C2 to the first group of cells
    names = Array(ws.range("C2").Value)
    group1Cells = Array("B10", "D10", "F10", "B15", "D15", "F15", "B20", "D20", "F20", "B25", "D25", "F25", "B30", "D30", "F30", _
                        "C9", "E9", "C14", "E14", "C19", "E19", "C24", "E24", "C29", "E29")
    
    For i = LBound(group1Cells) To UBound(group1Cells)
        name = Join(names, ", ")
        ws.range(group1Cells(i)).Value = name
    Next i

    ' Assign additional names from D2 to the second group of ranges
    additionalNames = Array(ws.range("D2").Value)
    group1Cells = Array("B9:F9", "B14:F14", "B19:F19", "B24:F24", "B29:F29")
    
    For i = LBound(group1Cells) To UBound(group1Cells)
        Set range = ws.range(group1Cells(i))
        For Each cell In range
            If IsEmpty(cell.Value) Then
                cell.Value = Join(additionalNames, ", ")
            Else
                cell.Value = cell.Value & ", " & Join(additionalNames, ", ")
            End If
        Next cell
    Next i

    ' Assign additional names from E2, E3, E4 to the third group of cells
    additionalNames3 = Array(ws.range("E2").Value, ws.range("E3").Value, ws.range("E4").Value)
    group1Cells = Array("B10", "C9", "D10", "E9", "F10", "B14", "C15", "D14", "E15", "F14", "B20", "C19", "D20", "E19", "F20", _
                        "B24", "C25", "D24", "E25", "F24", "B30", "C29", "D30", "E29", "F30")
    
    For i = LBound(group1Cells) To UBound(group1Cells)
        Set cell = ws.range(group1Cells(i))
        If IsEmpty(cell.Value) Then
            cell.Value = Join(additionalNames3, ", ")
        Else
            cell.Value = cell.Value & ", " & Join(additionalNames3, ", ")
        End If
    Next i

    ' Assign additional names from F2, F3, F4 to the fourth group of cells
    additionalNames2 = Array(ws.range("F2").Value, ws.range("F3").Value, ws.range("F4").Value)
    group1Cells = Array("B9", "C10", "D9", "E10", "F9", "B15", "C14", "D15", "E14", "F15", "B19", "C20", "D19", "E20", "F19", _
                        "B25", "C24", "D25", "E24", "F25", "B29", "C30", "D29", "E30", "F29")
    
    For i = LBound(group1Cells) To UBound(group1Cells)
        Set cell = ws.range(group1Cells(i))
        If IsEmpty(cell.Value) Then
            cell.Value = Join(additionalNames2, ", ")
        Else
            cell.Value = cell.Value & ", " & Join(additionalNames2, ", ")
        End If
    Next i
End Sub


