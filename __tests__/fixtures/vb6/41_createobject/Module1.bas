Attribute VB_Name = "Module1"
Option Explicit

Public Sub Run()
    Dim app As Object
    Set app = CreateObject("Excel.Application")
    app.Visible = True
End Sub
