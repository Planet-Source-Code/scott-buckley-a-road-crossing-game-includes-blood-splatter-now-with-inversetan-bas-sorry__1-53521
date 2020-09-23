Attribute VB_Name = "InverseTan"
Public Function ATan2(y, x)
    ATan2 = vbNull
    If x = 0 And y = 0 Then
        Exit Function
    ElseIf x = 0 And y < 0 Then
        ATan2 = pi / 2
    ElseIf x < 0 Then
        ATan2 = pi - Atn(y / x)
    ElseIf x = 0 And y > 0 Then
        ATan2 = -pi / 2
    ElseIf x > 0 Then
        ATan2 = 2 * pi - Atn(y / x)
    End If
End Function

