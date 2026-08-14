VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit

Public Function Compute() As Long
    Compute = 2
End Function

Public Sub Run()
    Dim v As Long
    v = Compute()
End Sub
