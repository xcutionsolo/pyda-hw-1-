Attribute VB_Name = "Module1"
Sub odd_even(x As Integer)
    x = InputBox("������� �����")
    If x Mod 2 = 0 Then
        MsgBox ("������")
    Else
        MsgBox ("�����")
    End If
End Sub

Sub newmacro()
    odd_even (2)
End Sub
