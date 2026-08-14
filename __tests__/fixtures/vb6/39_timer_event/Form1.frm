VERSION 5.00
Begin VB.Form Form1
   Caption         =   "Form1"
   ClientHeight    =   3195
   Begin VB.Timer tmrTick
      Interval        =   1000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_PredeclaredId = True
Option Explicit

Private Sub tmrTick_Timer()
    Poll
End Sub

Private Sub Poll()
End Sub
