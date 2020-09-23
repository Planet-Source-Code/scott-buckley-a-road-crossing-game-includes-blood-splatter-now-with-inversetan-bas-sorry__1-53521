VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   12135
   ClientLeft      =   1290
   ClientTop       =   1635
   ClientWidth     =   13680
   ForeColor       =   &H000000C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":247A
   ScaleHeight     =   12135
   ScaleWidth      =   13680
   Begin VB.Timer dudetimer 
      Interval        =   10
      Left            =   13080
      Top             =   1440
   End
   Begin VB.Timer movecars 
      Interval        =   30
      Left            =   13080
      Top             =   960
   End
   Begin VB.Label lblclsblood 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clean Up Blood"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label bstoggle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Blood Stains On"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblcarspeed 
      BackStyle       =   0  'Transparent
      Caption         =   "Car Speed:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblarmswing 
      BackStyle       =   0  'Transparent
      Caption         =   "Arm Swing:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblswing 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblswingup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblswingdown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblspeeddown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblspeedup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblspeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label deathtoggle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Death On"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label more 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10 More Cars"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape larm 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   11520
      Width           =   105
   End
   Begin VB.Shape rarm 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   11520
      Width           =   105
   End
   Begin VB.Shape dude 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   11520
      Width           =   195
   End
   Begin VB.Shape car 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   495
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub morecars(road As Integer, Dirp As Integer)
    Load car(car.Count)
    Cars = car.Count - 1
    Direction(Cars) = Dirp
    OnRoad(Cars) = road
    With car(Cars)
        .Visible = True
        .FillColor = RGB(Rnd * 155 + 100, Rnd * 155 + 100, Rnd * 155 + 100)
        .Left = Interval(Rnd * ScrWid, 495)
        .Top = PossY(road, Dirp)
    End With
End Sub

Private Sub dudetimer_Timer()
    CheckKeys
    MoveDude
    CalcArms
    DoArms
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown (KeyCode)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyUp (KeyCode)
End Sub

Private Sub Form_Load()
    doypos
    Restart
    Speed = 100
    Death = True
    ArmSwing = 20
End Sub

Private Sub Form_Resize()
    ScrWid = ScaleWidth
    ScrHei = ScaleHeight
End Sub

Private Sub deathtoggle_Click()
    If Death = True Then
        Death = False
        deathtoggle.Caption = "Death Off"
    Else
        Death = True
        deathtoggle.Caption = "Death On"
    End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub bstoggle_Click()
    If AutoRedraw = True Then
        Form1.AutoRedraw = False
        bstoggle.Caption = "Blood Stains Off"
    Else
        Form1.AutoRedraw = True
        bstoggle.Caption = "Blood Stains On"
    End If
End Sub

Private Sub lblclsblood_Click()
    If Form1.AutoRedraw = False Then
        Form1.AutoRedraw = True
        Form1.Cls
        Form1.AutoRedraw = False
    Else
        Form1.Cls
    End If
End Sub

Private Sub lblspeedup_Click()
    If Speed < 200 Then Speed = Speed + 10
    lblspeed.Caption = Speed
End Sub

Private Sub lblspeeddown_Click()
    If Speed > -200 Then Speed = Speed - 10
    lblspeed.Caption = Speed
End Sub

Private Sub lblspeedup_DblClick()
    lblspeedup_Click
End Sub

Private Sub lblspeeddown_DblClick()
    lblspeeddown_Click
End Sub

Private Sub lblswingdown_Click()
    If ArmSwing > -40 Then ArmSwing = ArmSwing - 10
    lblswing.Caption = ArmSwing
End Sub

Private Sub lblswingup_Click()
    If ArmSwing < 40 Then ArmSwing = ArmSwing + 10
    lblswing.Caption = ArmSwing
End Sub

Private Sub lblswingup_DblClick()
    lblswingup_Click
End Sub

Private Sub lblswingdown_DblClick()
    lblswingdown_Click
End Sub

Private Sub more_Click()
    For d = 0 To 9
        morecars Rnd * 7, Rnd * 1
    Next d
End Sub

Private Sub movecars_Timer()
    For x = 0 To Cars
        If Direction(x) = 0 Then
            car(x).Left = car(x).Left + Speed
        Else
            car(x).Left = car(x).Left - Speed
        End If
    
        RelocateCar (x)
        
        DoRect (x)
        
        DoInt (x)
    Next x
End Sub

Public Sub RelocateCar(ind As Integer)
    If car(ind).Left < -375 Then
        car(ind).Top = PossY(Round(Rnd * 7, 0), Opp(Direction(ind)))
        Direction(ind) = Opp(Direction(ind))
    End If
    
    If car(ind).Left > ScrWid Then
        car(ind).Top = PossY(Round(Rnd * 7, 0), Opp(Direction(ind)))
        Direction(ind) = Opp(Direction(ind))
    End If
End Sub

Public Sub DoRect(ind As Integer)
    With CarRect(ind)
        .Top = car(ind).Top
        .Left = car(ind).Left
        .Bottom = .Top + Hei
        .Right = .Left + Wid
    End With
End Sub

Public Sub DoInt(ind As Integer)
    If IntersectRect(IntersectArea, CarRect(ind), DudeRect) Then
        Splash cX, cY, ATan2(cY - (CarRect(ind).Top + 120), cX - (CarRect(ind).Left + 182)) * AngInv, 1000, 30
        If Death = True Then Restart
    End If
End Sub

Public Sub MoveDude()
    dude.Left = dude.Left + Vel * Cos(Ang)
    dude.Top = dude.Top - Vel * Sin(Ang)
    cX = dude.Left + 100
    cY = dude.Top + 100
    
    With DudeRect
        .Top = dude.Top
        .Left = dude.Left
        .Bottom = .Top + 200
        .Right = .Left + 200
    End With
End Sub

Public Sub DoArms()
    ArmPl = ArmPl + 4 * RadInv
    If ArmPl = 360 Then arpl = 0
    swingvel = ((Vel + 0.0000001) / MaxSpeed) * ArmSwing
    temp = RadInv * swingvel * Sin(ArmPl)
    larm.Left = cX - 150 * Cos(260 * RadInv - Ang + temp) - 50
    rarm.Left = cX + 150 * Cos(280 * RadInv - Ang + temp) - 50
    larm.Top = cY - 150 * Sin(260 * RadInv - Ang + temp) - 50
    rarm.Top = cY + 150 * Sin(280 * RadInv - Ang + temp) - 50
End Sub

Public Sub Splash(x As Single, y As Single, Angle As Integer, Length As Integer, Threshhold As Integer)
    For p = 0 To 30
        tmprnd = Rnd * Threshhold * 2 - Threshhold
        tmpang = Angle + tmprnd
        rt = Rnd * 100
        ForeColor = RGB(Rnd * 100 + 155, rt, rt)
        Line (x, y)-(x + (Rnd * Length) * Sin(RadInv * tmpang), y + (Rnd * Length) * Cos(RadInv * tmpang))
    Next p
End Sub

Public Sub Restart()
    dude.Left = 5760
    dude.Top = 11520
    Ang = 90 * RadInv
    Vel = 0
End Sub
