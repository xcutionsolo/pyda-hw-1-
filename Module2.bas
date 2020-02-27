Attribute VB_Name = "Module2"
Public Function mult_odd_even(ByVal x As Long, ByVal y As Long)
mult_odd_even = (x * y) Mod 2
End Function

Sub newc()
    Dim a As Long
    Dim b As Long
    Dim c As Long
    a = InputBox("¬ведите 1ое число")
    b = InputBox("¬ведите второе число")
    c = mult_odd_even(a, b)
    If c = 0 Then
        MsgBox ("четное")
    Else
        MsgBox ("нечет")
    End If


End Sub


Sub mass()
Dim arr(0 To 8) As Integer
For i = 0 To 8
    arr(i) = i * 3
    Next i
b = arr(0)
For a = 0 To 8
    If arr(a) > b Then
    b = arr(a)
    End If
    Next a
MsgBox (b)
    
    


End Sub
