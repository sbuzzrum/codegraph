VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientWidth     =   4680
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mReader As Class1

Private Sub Form_Load()
    Set mReader = New Class1
End Sub

Private Sub mReader_ItemRead(ByVal Code As String)
    Debug.Print Code
End Sub
