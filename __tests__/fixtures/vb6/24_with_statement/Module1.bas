Attribute VB_Name = "Module1"
Option Explicit

Public Sub UseWith()
    Dim c As New Class1
    With c
        .Value = 3
        .Compute 2
    End With
End Sub
