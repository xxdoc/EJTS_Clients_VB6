VERSION 5.00
Begin VB.Form tabPullFiles 
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
   Begin VB.ListBox lstSort 
      Height          =   300
      IntegralHeight  =   0   'False
      ItemData        =   "tabPullFiles.frx":0000
      Left            =   2160
      List            =   "tabPullFiles.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PLFL_pctFrame 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "tabPullFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ITab
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private Const PLFL_OffsetX = 8
Private Const PLFL_OffsetY = 8
Private Const PLFL_RCX = 90
Private Const PLFL_NameX = 150
Private Const PLFL_LineHeight = 50
Private Const PLFL_FontSize = 40
Public PLFL_Font&

Private Sub Form_Load()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

Private Function ITab_CreateGDIObjects() As Boolean
'errheader>
Const PROC_NAME = "tabPullFiles" & "." & "ITab_CreateGDIObjects": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

PLFL_Font = CreateFont2(PLFL_pctFrame.hdc, "Arial", PLFL_FontSize, False, False, False, False)

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Function ITab_InitializeAfterDBLoad() As Boolean
End Function

Private Sub ITab_AfterTabShown()
'errheader>
Const PROC_NAME = "tabPullFiles" & "." & "ITab_AfterTabShown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, b&, ai&, ci&, vi&, ty&, t$, pcthdc&, td As Date

'Clear
PLFL_pctFrame.Cls

'Find tomorrow's date
t$ = DB_GetSetting(ActiveDBInstance, "GLOBAL_PullFilesWeekdaysToSkip")
b = Date + 1
For a = b To b + 7      'Don't do an infinite loop here
    If InStr(t$, WeekdayName(Weekday(a), True)) = 0 Then
        td = a
        Exit For
    End If
Next a
If td = 0 Then ERR_COUNT = ERR_COUNT + 1: GoTo CLEAN_UP

'Sort appointments by time slot, the way they are shown on the schedule
lstSort.Clear
For ai = 0 To ActiveDBInstance.Appointments_Count - 1
    With ActiveDBInstance.Appointments(ai)
        If .ApptDate = td Then
            lstSort.AddItem Format$(.ApptTimeSlot, "000")
            lstSort.ItemData(lstSort.NewIndex) = ai
        End If
    End With
Next ai

'Draw
pcthdc = PLFL_pctFrame.hdc
SelectObject pcthdc, PLFL_Font
If lstSort.ListCount = 0 Then
    t$ = "(No apppointments on " & Format$(td, "ddd, m/d") & ")"
    SetTextColor pcthdc, vbBlack
    TextOut pcthdc, PLFL_OffsetX + PLFL_NameX, ty, t$, Len(t$)
Else
    For a = 0 To lstSort.ListCount - 1
        ai = lstSort.ItemData(a)
        For b = 0 To ActiveDBInstance.Appointments(ai).ClientID_Count - 1
            ci = DB_FindClientIndex(ActiveDBInstance, ActiveDBInstance.Appointments(ai).ClientIDs(b))
            With ActiveDBInstance.Clients(ci).c
                ty = PLFL_OffsetY + (vi * PLFL_LineHeight)
                'New client / NNTF
                If Flag_IsSet(.Flags, NewClient) Then
                    t$ = "N"
                    SetTextColor pcthdc, &H7FFF
                ElseIf Flag_IsSet(.LastYear_Flags, NoNeedToFile) Then
                    t$ = "NF"
                    SetTextColor pcthdc, &H969696
                Else
                    t$ = Chr$(&HB7&)
                    SetTextColor pcthdc, &HD2D2D2
                End If
                TextOut pcthdc, PLFL_OffsetX, ty, t$, Len(t$)
                'Reminder call
                If b = 0 Then
                    If Flag_IsSet(ActiveDBInstance.Appointments(ai).Flags, ReminderCall) Then
                        t$ = "R"
                        If Flag_IsSet(ActiveDBInstance.Appointments(ai).Flags, Called) Then
                            SetTextColor pcthdc, &HA000&
                        Else
                            SetTextColor pcthdc, &HFF
                        End If
                    Else
                        t$ = Chr$(&HB7&)
                        SetTextColor pcthdc, &HD2D2D2
                    End If
                    TextOut pcthdc, PLFL_OffsetX + PLFL_RCX, ty, t$, Len(t$)
                End If
                'Name
                If b = 0 Then t$ = "" Else t$ = "+ "
                t$ = t$ & FormatClientName(fPullFiles, ActiveDBInstance.Clients(ci).c)
                SetTextColor pcthdc, vbBlack
                TextOut pcthdc, PLFL_OffsetX + PLFL_NameX, ty, t$, Len(t$)
                vi = vi + 1
            End With
        Next b
    Next a
End If
PLFL_pctFrame.Refresh

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
Const PROC_NAME = "tabPullFiles" & "." & "ITab_SetDefaultFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetFocusWithoutErr PLFL_pctFrame

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
'errheader>
Const PROC_NAME = "tabPullFiles" & "." & "ITab_DestroyGDIObjects": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

DeleteObject PLFL_Font

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Sub Form_Resize()
'errheader>
On Error Resume Next        'ALL ERRORS WILL BE IGNORED IN THIS PROCEDURE
'<errheader

PLFL_pctFrame.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
ITab_AfterTabShown

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabPullFiles" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabPullFiles" & "." & "Form_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabPullFiles" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

