Attribute VB_Name = "Module1"
Option Explicit

Public Sub UseClass()
    Dim c As Class1
    Set c = New Class1
    Call c.Compute(1)
    Dim v As Long
    v = c.GetValue()
End Sub
