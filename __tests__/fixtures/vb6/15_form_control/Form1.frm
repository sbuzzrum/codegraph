VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientWidth     =   4680
   Begin VB.Frame fraMain
      Caption         =   "Group"
      Begin VB.CommandButton cmdOk
         Caption         =   "OK"
      End
   End
   Begin VB.TextBox txtName
      Text            =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Debug.Print txtName.Text
End Sub
