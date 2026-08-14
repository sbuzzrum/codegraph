Attribute VB_Name = "Module1"
Option Explicit

#Const DEBUG_MODE = 1

Public Sub Run()
#If DEBUG_MODE Then
    LogDebug
#Else
    LogRelease
#End If
End Sub

Public Sub LogDebug()
End Sub

Public Sub LogRelease()
End Sub
