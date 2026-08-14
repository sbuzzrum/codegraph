VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientWidth     =   4680
   Begin VB.CommandButton cmdAction
      Caption         =   "One"
      Index           =   0
   End
   Begin VB.CommandButton cmdAction
      Caption         =   "Two"
      Index           =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)
    Debug.Print Index
End Sub
