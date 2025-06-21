
Sub CalculateSLA()

    Dim lastRow As Long
    Dim i As Long
    Dim resolutionTime As Double
    Dim slaLimit As Double

    lastRow = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        If Cells(i, 2).Value <> "" And Cells(i, 3).Value <> "" Then
            resolutionTime = (Cells(i, 3).Value - Cells(i, 2).Value) * 24
            Cells(i, 6).Value = resolutionTime

            Select Case Cells(i, 4).Value
                Case "Critical": slaLimit = 2
                Case "High": slaLimit = 4
                Case "Medium": slaLimit = 6
                Case "Low": slaLimit = 8
                Case Else: slaLimit = 5
            End Select

            If resolutionTime <= slaLimit Then
                Cells(i, 7).Value = "Within SLA"
            Else
                Cells(i, 7).Value = "Breached"
            End If
        End If
    Next i

End Sub
