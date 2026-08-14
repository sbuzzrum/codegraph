VERSION 5.00
Begin VB.UserControl MyControl
   ClientHeight    =   1200
End
Attribute VB_Name = "MyControl"
Attribute VB_Exposed = True
Option Explicit

Public Event Changed()

Public Sub Clear()
End Sub

Public Property Get Text() As String
End Property
