Attribute VB_Name = "ModA"
Option Explicit

Private Function Compute(ByVal n As Long) As Long
    Compute = n
End Function

Public Sub RunA()
    Dim r As Long
    r = Compute(1)
End Sub
