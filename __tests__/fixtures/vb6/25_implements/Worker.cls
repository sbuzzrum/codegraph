VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
END
Attribute VB_Name = "Worker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit

Implements IWorker

Private Sub IWorker_Run()
    Debug.Print "running"
End Sub
