Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Run()
    Dim t As Long
    t = GetTickCount()
    Sleep 10
End Sub
