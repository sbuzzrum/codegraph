VERSION 5.00
Object = "{67397AA3-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1
   Caption         =   "Form1"
   Begin MSAdodcLib.Adodc Adodc1
      Left            =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_PredeclaredId = True
Option Explicit

Private Sub Refresh2()
    Adodc1.Recordset.MoveNext
End Sub
