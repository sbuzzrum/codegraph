VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   ClientHeight    =   3195
   Begin VB.CommandButton cmdSave
      Caption         =   "Save"
   End
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_PredeclaredId = True
Option Explicit

Private Sub cmdSave_Click()
    SaveIt
End Sub

Private Sub cmdCancel_Click()
    CancelIt
End Sub

Private Sub SaveIt()
End Sub

Private Sub CancelIt()
End Sub
