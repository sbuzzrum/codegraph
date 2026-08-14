VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   Begin VB.TextBox txtName
      Text            =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_PredeclaredId = True
Option Explicit

Private Sub Apply()
    txtName.Text = "hello"
    txtName.SetFocus
End Sub
