VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobalSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "CGlobalSettings"

'By the time we are exiting the program for real, Sub Main() will already have ended, and frmStart has already unloaded
'So we use this as a tricky way of ensuring we can run some last-minute code before quitting
'EHT=Standard
Private Sub Class_Terminate()
On Error GoTo ERR_HANDLER

Main_Unload

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Class_Terminate", Err
End Sub
