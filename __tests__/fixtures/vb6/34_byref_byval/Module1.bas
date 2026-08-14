Attribute VB_Name = "Module1"
Option Explicit

Public Sub Swap2(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a
    a = b
    b = t
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function

Public Sub Run()
    Dim x As Long, y As Long
    Swap2 x, y
    x = Twice(y)
End Sub
