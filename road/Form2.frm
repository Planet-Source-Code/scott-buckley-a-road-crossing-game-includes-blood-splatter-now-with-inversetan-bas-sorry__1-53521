VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   12135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   12135
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "More"
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   12840
      Top             =   240
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   0
      X1              =   0
      X2              =   13680
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   0
      Left            =   0
      Top             =   960
      Width           =   13695
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   7
      X1              =   0
      X2              =   13680
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Shape car 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   375
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   6
      X1              =   0
      X2              =   13680
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   5
      X1              =   0
      X2              =   13680
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   4
      X1              =   0
      X2              =   13680
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   3
      X1              =   0
      X2              =   13680
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   2
      X1              =   0
      X2              =   13680
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line linerep 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      Index           =   1
      X1              =   0
      X2              =   13680
      Y1              =   10680
      Y2              =   10680
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   7
      Left            =   0
      Top             =   3600
      Width           =   13695
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   6
      Left            =   0
      Top             =   2280
      Width           =   13695
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   5
      Left            =   0
      Top             =   4920
      Width           =   13695
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   4
      Left            =   0
      Top             =   6240
      Width           =   13695
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   3
      Left            =   0
      Top             =   7560
      Width           =   13695
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   2
      Left            =   0
      Top             =   8880
      Width           =   13695
   End
   Begin VB.Shape roadrep 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Index           =   1
      Left            =   0
      Top             =   10200
      Width           =   13695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub morecars()
    Load car(car.Count)
    With car(car.Count - 1)
        .Visible = True
        .FillColor = RGB(Rnd * 155 + 100, Rnd * 155 + 100, Rnd * 155 + 100)
        .Left = Rnd * ScrWid
        .Top = Rnd * ScrHei
    End With
End Sub

Private Sub Command1_Click()
    morecars
End Sub

Private Sub Form_Resize()
    ScrWid = Width - 375
    ScrHei = Height - 255
End Sub
