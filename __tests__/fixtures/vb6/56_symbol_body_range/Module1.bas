Attribute VB_Name = "Module1"
Option Explicit

Public Type TPoint
    X As Long
    Y As Long
End Type

Public Enum ColorKind
    ckRed = 1
    ckBlue = 2
End Enum

Public Sub Run()
    Dim total As Long
    total = 1
    total = total + 1
    Debug.Print total
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
