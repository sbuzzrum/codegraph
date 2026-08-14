Attribute VB_Name = "Module1"
Option Explicit

Public Sub Log(ByVal Message As String, Optional ByVal Level As Long = 1)
    Debug.Print Message, Level
End Sub

Public Sub Run()
    Log "hi"
    Log "hi", 2
    Log Message:="hi", Level:=3
End Sub
