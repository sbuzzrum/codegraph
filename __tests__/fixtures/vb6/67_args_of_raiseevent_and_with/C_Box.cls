VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Box"
Attribute VB_Exposed = True
Option Explicit

Public Event Changed(ByVal Value As Long)

Public Function Compute2(ByVal n As Long) As Long
    Compute2 = n
End Function

Public Function Child2() As C_Box
End Function

Public Sub Touch2()
End Sub

Public Sub Fire()
    RaiseEvent Changed(Compute2(1))
End Sub

Public Sub Walk()
    With Child2()
        .Touch2
    End With
End Sub
