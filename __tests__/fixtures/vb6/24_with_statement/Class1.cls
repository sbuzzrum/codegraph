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

Public Sub Compute(ByVal n As Long)
    Debug.Print n
End Sub

Public Property Let Value(ByVal v As Long)
    Debug.Print v
End Property
