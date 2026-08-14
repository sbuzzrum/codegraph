Attribute VB_Name = "Module1"
Option Explicit

Public Sub Run()
    Dim o As Object
    Set o = CreateObject("Some.Server")
    Call o.DoStuff(1)
End Sub
