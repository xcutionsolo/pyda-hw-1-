Attribute VB_Name = "Module3"
Sub hikan()
    Dim i As Integer, j As Integer
    Sheets("Лист3").Activate
    Randomize
    For i = 1 To 20 Step 2
        For j = 1 To 20 Step 2
            Cells(i, j).Value = Abs(Rnd * 100 - 50)
            Next j
        Next i
End Sub

Sub delrand()
    Dim i As Integer, j As Integer
    Sheets("Лист3").Activate
    Randomize
    For i = 1 To 20
        For j = 1 To 20
        Cells(i, j).Delete
End Sub
