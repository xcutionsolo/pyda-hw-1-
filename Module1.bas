Attribute VB_Name = "Module1"
Sub odd_even(x As Integer)
    x = InputBox("¬ведите число")
    If x Mod 2 = 0 Then
        MsgBox ("четное")
    Else
        MsgBox ("нечет")
    End If
End Sub

Sub newmacro()
    odd_even (2)
End Sub
