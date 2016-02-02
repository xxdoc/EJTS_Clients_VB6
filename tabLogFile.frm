VERSION 5.00
Begin VB.Form tabLogFile 
   BackColor       =   &H000080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13425
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   895
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLog 
      Height          =   540
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "tabLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabLogFile"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private ComputerName$

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=None
Private Function ITab_CreateGDIObjects() As Boolean
End Function

'EHT=Cleanup1
Private Function ITab_InitializeAfterDBLoad() As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim l&
l = 100
ComputerName$ = Space$(l)
GetComputerName ComputerName$, l
If l > 7 Then l = 7
ComputerName$ = Left$(ComputerName$, l)

Dim t$
lstLog.Clear
If FileExists(ActiveDBInstance.FullPath_Log) Then
    Dim fh As CMNMOD_CFileHandler
    Set fh = OpenFile(ActiveDBInstance.FullPath_Log, mLineByLine_Input)
    Do Until fh.EndOfFile
        t$ = fh.ReadLine
        lstLog.AddItem t$
    Loop
    fh.CloseFile: Set fh = Nothing
    lstLog.TopIndex = lstLog.ListCount - 1
End If

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err, INCLEANUP: Resume CLEANUP
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

lstLog.ListIndex = -1

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr lstLog

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SetDefaultFocus", Err
End Sub

'EHT=None
Private Function ITab_SaveSettingsToDBBeforeClose() As Boolean
End Function

'EHT=None
Private Function ITab_DestroyGDIObjects() As Boolean
End Function

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

lstLog.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

'EHT=Standard
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=Standard
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub     'Pass it to the parent form first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyUp", Err
End Sub

'EHT=Standard
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyPress KeyAscii: If KeyAscii = 0 Then Exit Sub       'Pass it to the parent form first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyPress", Err
End Sub

'EHT=Standard
Private Sub lstLog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER

'Select item under mouse
Dim i&
i = SendMessage(lstLog.hwnd, LB_ITEMFROMPOINT, 0, (X / Screen.TwipsPerPixelX) + ((Y / Screen.TwipsPerPixelY) * &H10000))
If i > &HFFFF& Then
    lstLog.ListIndex = -1
Else
    i = (i And &HFFFF&)
    If Button = vbRightButton Then
        lstLog.ListIndex = i    'Listbox only does this for left click on a valid item
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstLog_MouseDown", Err
End Sub

'EHT=Cleanup1
Public Sub WriteLine(ByVal t$)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

If Not ActiveDBInstance.IsWriteable Then Exit Sub

Dim ts As Date, d As Long
Static ld As Long
ts = Now
d = Int(ts)
t$ = Format$(ts, "yyyy-mm-dd hh:mma/p") & vbTab & ComputerName & vbTab & t$
Dim fh As CMNMOD_CFileHandler
Set fh = OpenFile(ActiveDBInstance.FullPath_Log, mLineByLine_Append)
If d <> ld Then
    fh.WriteLine ""   'Print a blank line to separate days (also if first log of session)
    lstLog.AddItem ""
    ld = d
End If
fh.WriteLine t$
fh.CloseFile: Set fh = Nothing
lstLog.AddItem t$
lstLog.TopIndex = lstLog.ListCount - 1

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "WriteLine", Err, INCLEANUP: Resume CLEANUP
End Sub

