Attribute VB_Name = "modIntersectRect"
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
End Type

Public IntersectArea As RECT
