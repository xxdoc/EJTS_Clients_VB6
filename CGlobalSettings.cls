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

'By the time we are exiting the program for real, Sub Main() will already have ended, and frmStart has already unloaded
'So we use this as a tricky way of ensuring we can run some last-minute code before quitting
Private Sub Class_Terminate()
Main_Unload
End Sub
