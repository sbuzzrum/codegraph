Attribute VB_Name = "Module1"
Option Explicit

Public Sub Risky()
    On Error GoTo ErrHandler
    Cleanup
    Exit Sub
ErrHandler:
    Resume Next
End Sub

Public Sub Cleanup()
End Sub
