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

Public Event ItemRead(ByVal Code As String)

Public Sub ReadBarcode(ByVal Code As String)
    RaiseEvent ItemRead(Code)
End Sub
