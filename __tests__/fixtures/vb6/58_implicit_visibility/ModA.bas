Attribute VB_Name = "ModA"
Option Explicit

Dim mModuleLevel As Long

Sub NoModifier()
    Debug.Print "implicitly public"
End Sub

Function AlsoPublic() As Long
    AlsoPublic = 1
End Function
