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
Implements ITab

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to frmMain first, Exit if form cancelled the event
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
frmMain.Form_KeyPress KeyAscii: If KeyAscii = 0 Then Exit Sub       'Pass it to frmMain first, Exit if form cancelled the event
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub     'Pass it to frmMain first, Exit if form cancelled the event
End Sub

Private Sub Form_Resize()
On Error Resume Next
End Sub

Private Function ITab_Initialize() As Boolean
'
End Function

Private Function ITab_LoadSettings() As Boolean
'
End Function

Private Sub ITab_AfterTabShown()
'
End Sub

Private Sub ITab_SetDefaultFocus()
On Error Resume Next
End Sub

Private Sub ITab_AfterTabHidden()
'
End Sub

Private Sub ITab_CodeRequestingRefresh()
'
End Sub

Private Function ITab_CommitDataToDB() As Boolean
'
End Function

Private Function ITab_CleanUp() As Boolean
'
End Function
