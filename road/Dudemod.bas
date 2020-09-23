Attribute VB_Name = "Dudemod"
Public cX As Single
Public cY As Single

Public Const ArmRad As Integer = 150

Public Const MinSpeed As Double = -10
Public Const MaxSpeed As Integer = 30

Public Const SlowDown As Double = 0.95

Public Const Pi As Double = 3.14159265358979
Public Const RadInv As Double = Pi / 180
Public Const AngInv As Double = 180 / Pi

Public ArmPl As Single
Public ArmSwing As Integer

Public Ang As Double
Public Vel As Double

Public LX As Integer
Public LY As Integer

Public RX As Integer
Public RY As Integer

Public DudeRect As RECT

Public Sub CalcArms()
    LX = cX - DudeRad * Sin(Ang)
    LY = cY - DudeRad * Cos(Ang)
    RX = cX + cX - LX
    RY = cY + cY - LY
End Sub
