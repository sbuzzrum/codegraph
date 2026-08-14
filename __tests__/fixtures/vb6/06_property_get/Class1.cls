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

Private mValue As Long

Public Property Get Value() As Long
    Value = mValue
End Property
