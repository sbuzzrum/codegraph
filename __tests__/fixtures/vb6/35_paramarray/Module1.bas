Attribute VB_Name = "Module1"
Option Explicit

Public Sub LogAll(ParamArray Items() As Variant)
    Debug.Print UBound(Items)
End Sub

Public Sub Run()
    LogAll "a", "b", "c"
End Sub
