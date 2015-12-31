VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function CreateGDIObjects() As Boolean
'Create any fonts, brushes, etc. that must be destroyed later in the DestroyGDIObject() call (Database is not loaded yet)
End Function

Function InitializeAfterDBLoad() As Boolean
'DB is now loaded and accessible
End Function

Sub AfterTabShown()
'Primary drawing/update procedure
End Sub

Sub SetDefaultFocus()
'Tab-form should focus the default control when this is called
End Sub

Function SaveSettingsToDBBeforeClose() As Boolean
'Write data to database so it can be saved to file
'(this will only get called if the database was actually loaded and the changed flag was triggered)
End Function

Function DestroyGDIObjects() As Boolean
'Destroy any fonts, brushes, etc. created during CreateGDIObjects()
End Function
