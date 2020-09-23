Attribute VB_Name = "Other"
Public Direction(1000) As Integer
Public OnRoad(1000) As Integer
Public PossY(7, 1) As Single
Public CarRect(1000) As RECT

Public ScrWid As Single
Public ScrHei As Single

Public Cars As Integer

Public Speed As Integer
Public Const Wid As Integer = 375
Public Const Hei As Integer = 255

Public Death As Boolean

Public Sub doypos()
    For x = 0 To 7
        PossY(x, 0) = 1080 + x * 1320
        PossY(x, 1) = PossY(x, 0) + 495
    Next x
End Sub

Public Function Opp(Inp As Integer)
    If Inp = 1 Then Opp = 0 Else Opp = 1
End Function

Public Function Interval(Inp As Variant, Intrvl As Variant)
    Interval = Intrvl * Round(Inp / Intrvl, 0)
End Function
