VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMousePointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ModName = "clsMousePointer"

Private nOldMousPointer    As Long

Private Sub e()
    On Error GoTo ERRHANDLER
    '// --- Show the Hourglass ---------------------------------------------------------------//
    nOldMousPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    Exit Sub
    
ERRHANDLER:
    MsgBox "Unexpected error in procedure: Class_Initialize" & vbCrLf & _
        "Error #" & Err.Number & ": " & Err.Description, _
        vbCritical + vbOKOnly, App.Title
End Sub

Private Sub Class_Terminate()
    On Error GoTo ERRHANDLER
    '// --- Show previous pointer ------------------------------------------------------------//
    Screen.MousePointer = nOldMousPointer
    Exit Sub
    
ERRHANDLER:
    MsgBox "Unexpected error in procedure: Class_Terminate" & vbCrLf & _
        "Error #" & Err.Number & ": " & Err.Description, _
        vbCritical + vbOKOnly, App.Title
End Sub




