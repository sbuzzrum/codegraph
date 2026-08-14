Attribute VB_Name = "Module1"
Option Explicit

Public Sub Tick()
    Static callCount As Long
    callCount = callCount + 1
End Sub
