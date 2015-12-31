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
Implements ITab
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private ComputerName$

Private Sub Form_Load()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

Private Function ITab_CreateGDIObjects() As Boolean
End Function

Private Function ITab_InitializeAfterDBLoad() As Boolean
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "ITab_InitializeAfterDBLoad": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Sub ITab_AfterTabShown()
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "ITab_AfterTabShown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

lstLog.ListIndex = -1

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub ITab_SetDefaultFocus()
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "ITab_SetDefaultFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetFocusWithoutErr lstLog

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Function ITab_SaveSettingsToDBBeforeClose() As Boolean
End Function

Private Function ITab_DestroyGDIObjects() As Boolean
End Function

Private Sub Form_Resize()
'errheader>
On Error Resume Next        'ALL ERRORS WILL BE IGNORED IN THIS PROCEDURE
'<errheader

lstLog.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "Form_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub     'Pass it to the parent form first, Exit if form cancelled the event

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

frmMain.Form_KeyPress KeyAscii: If KeyAscii = 0 Then Exit Sub       'Pass it to the parent form first, Exit if form cancelled the event

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lstLog_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "lstLog_MouseDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Select item under mouse
Dim i&
i = SendMessage(lstLog.hwnd, LB_ITEMFROMPOINT, 0, (X / Screen.TwipsPerPixelX) + ((y / Screen.TwipsPerPixelY) * &H10000))
If i > &HFFFF& Then
    lstLog.ListIndex = -1
Else
    i = (i And &HFFFF&)
    If Button = vbRightButton Then
        lstLog.ListIndex = i    'Listbox only does this for left click on a valid item
    End If
End If

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub WriteLine(ByVal t$)
'errheader>
Const PROC_NAME = "tabLogFile" & "." & "WriteLine": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

