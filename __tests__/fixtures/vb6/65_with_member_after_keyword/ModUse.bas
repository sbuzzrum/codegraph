Attribute VB_Name = "ModUse"
Option Explicit

Public Sub Run()
    Dim d As New C_Data
    With d
        If .Count2 > 0 Then
            .Reset2
        End If
    End With
End Sub
