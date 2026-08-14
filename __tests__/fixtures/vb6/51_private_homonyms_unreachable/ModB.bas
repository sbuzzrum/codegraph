Attribute VB_Name = "ModB"
Option Explicit

Private Function Compute(ByVal n As Long) As Long
    Compute = n * 2
End Function

Public Sub RunB()
    Dim r As Long
    r = Compute(2)
End Sub
