VERSION 5.00
Begin VB.UserControl MyControl
   ClientHeight    =   1200
   ClientWidth     =   1800
   Begin VB.Label lblCaption
      Caption         =   "x"
   End
End
Attribute VB_Name = "MyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Changed(ByVal NewValue As String)

Private mText As String

Public Property Get Text() As String
    Text = mText
End Property

Public Sub Clear()
    mText = ""
End Sub
