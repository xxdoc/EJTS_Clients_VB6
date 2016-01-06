VERSION 5.00
Begin VB.Form tab___Template 
   BorderStyle     =   0  'None
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13425
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   895
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "tab___Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tab___Template"
Implements ITab

'EHT=Standard
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to frmMain first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=Standard
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyPress KeyAscii: If KeyAscii = 0 Then Exit Sub       'Pass it to frmMain first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyPress", Err
End Sub

'EHT=Standard
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub     'Pass it to frmMain first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyUp", Err
End Sub

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

On Error Resume Next
End Sub

'EHT=Standard
Private Function ITab_Initialize() As Boolean
On Error GoTo ERR_HANDLER

'

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_Initialize", Err
End Function

'EHT=Standard
Private Function ITab_LoadSettings() As Boolean
On Error GoTo ERR_HANDLER

'

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_LoadSettings", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

'

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=ResumeNext
Private Sub ITab_SetDefaultFocus()
On Error Resume Next

On Error Resume Next
End Sub

'EHT=Standard
Private Sub ITab_AfterTabHidden()
On Error GoTo ERR_HANDLER

'

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabHidden", Err
End Sub

'EHT=Standard
Private Sub ITab_CodeRequestingRefresh()
On Error GoTo ERR_HANDLER

'

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CodeRequestingRefresh", Err
End Sub

'EHT=Standard
Private Function ITab_CommitDataToDB() As Boolean
On Error GoTo ERR_HANDLER

'

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CommitDataToDB", Err
End Function

'EHT=Standard
Private Function ITab_CleanUp() As Boolean
On Error GoTo ERR_HANDLER

'

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CleanUp", Err
End Function
