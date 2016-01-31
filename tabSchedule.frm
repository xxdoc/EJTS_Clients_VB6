VERSION 5.00
Begin VB.Form tabSchedule 
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
   Begin VB.PictureBox pctSchedule 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H80000012&
      Height          =   1935
      Left            =   960
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   840
      Width           =   4815
      Begin VB.Label lblApptSelection 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Move (hold Ctrl for copy)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2775
      End
      Begin VB.Shape shpApptSelection 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image imgCurTime 
         Height          =   105
         Left            =   360
         Picture         =   "tabSchedule.frx":0000
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgReminderCall 
         Height          =   180
         Index           =   1
         Left            =   840
         Picture         =   "tabSchedule.frx":0073
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image imgReminderCall 
         Height          =   180
         Index           =   0
         Left            =   600
         Picture         =   "tabSchedule.frx":00EC
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin VB.Timer tmrFlashAppt 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   600
   End
   Begin VB.Menu menAppt 
      Caption         =   "Appt"
      Visible         =   0   'False
      Begin VB.Menu menAppt_Title 
         Caption         =   "======= 9:00 AM ======="
         Enabled         =   0   'False
      End
      Begin VB.Menu menAppt_Sep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menApptEdit 
         Caption         =   "&Edit Appt..."
      End
      Begin VB.Menu menApptReschedule 
         Caption         =   "&Reschedule"
      End
      Begin VB.Menu menApptCancelDelete 
         Caption         =   "&Cancel / &Delete"
      End
      Begin VB.Menu menApptScheduleFromThis 
         Caption         =   "New appointment using these client names"
      End
      Begin VB.Menu menApptMarkReminderCall 
         Caption         =   "Re&minder Call"
      End
      Begin VB.Menu menApptMarkCalled 
         Caption         =   "Ca&lled"
      End
      Begin VB.Menu menApptMarkDidntHappen 
         Caption         =   "Didn't Happe&n"
      End
      Begin VB.Menu menApptCLItem 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menSlot 
      Caption         =   "Slot"
      Visible         =   0   'False
      Begin VB.Menu menSlot_Title 
         Caption         =   "=== 9:00 AM ==="
         Enabled         =   0   'False
      End
      Begin VB.Menu menSlot_Sep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menSlotMarkDefault 
         Caption         =   "&Default (template)"
      End
      Begin VB.Menu menSlot_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menSlotMarkAvailable 
         Caption         =   "&Available"
      End
      Begin VB.Menu menSlotMarkReserved 
         Caption         =   "&Reserved"
      End
      Begin VB.Menu menSlotMarkMealBreak 
         Caption         =   "&Meal break"
      End
      Begin VB.Menu menSlotCreateNonClient 
         Caption         =   "Non-&client item..."
      End
   End
End
Attribute VB_Name = "tabSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabSchedule"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

#Const LRTB = True

Public ScheduleMode As enumScheduleMode
Private ApptIDBeingRescheduled As Long
Private ClickedDayIndex As Long
Private ClickedDate As Long
Private ClickedTimeslot As Long
Private ClickedApptIndex As Long
Private DoubleClickAllowed As Boolean
Private pctScheduleHdc&
Public ViewStartDate As Long
Public TodaysDayIndex&
Private ScheduleDayPositions(6) As RECT
Private FontTitle As Long
Private FontSubtitle As Long
Private FontTimesOnSlot As Long
Private FontTimesOffSlot As Long
Private FontApptPrimary As Long
Private FontApptSecondary As Long
Private FontApptMinutes As Long
Private Const ColorTimeText_Empty = vbBlack
Private Const ColorTimeText_OnSlot = vbBlack
Private Const ColorTimeText_OffSlot = vbRed
Private Const ColorSlotMealBreak = &HC0C0C0
Private Const ColorSlotReserved = &HD8D8D8
Private ColorProfilesDay()   '(index,color) 0=titletext, 1=titlebg, 2=border, 3=apptbg
Private ColorProfilesAppt()  '(index,color) 0=text,      1=bg,      2=border, 3=postedcolor
Private Const MarginTop = 1 '3
Private Const MarginLeft = 10 '3
Private Const DaySpacingX = 2 '9
Private Const DaySpacingY = 2 '9
Private Const DayMarginTop = 3
Private Const DayMarginBottom = 3
Private Const DayMarginLeft = 3
Private Const DayMarginRight = 3
Private Const DayTitleHeight = 43
Private Const DayFirstSlotOffsetY = DayTitleHeight + DayMarginTop
Private Const DayTimesOffsetX = DayMarginLeft + 35
Private Const DayApptsOffsetX = DayTimesOffsetX + 13
Private DayApptSlotHeight&    '12 for KNW, 17 for all others
Private Const DayWidth = 320
Private DayHeight&            'DayFirstSlotOffsetY + (DayApptSlotHeight * Appointment_NumSlots) + 1 + DayMarginBottom + 1
Private LastShapeStyle As ScheduleShapeStyle
Private ScheduleTemplate(1 To 3, 6, Appointment_NumSlotsUB) As Long     '3D array
Private lastMouseMoveX!
Private lastMouseMoveY!

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Standard
Private Function ITab_CreateGDIObjects() As Boolean
On Error GoTo ERR_HANDLER

FontTitle = CreateFont2(pctSchedule.hdc, "Arial", 14, False, False, False, False)
FontSubtitle = CreateFont2(pctSchedule.hdc, "Arial", 12, False, False, False, False)
FontTimesOnSlot = CreateFont2(pctSchedule.hdc, "Arial", 8, False, False, False, False)
FontTimesOffSlot = CreateFont2(pctSchedule.hdc, "Arial", 8, True, True, False, False)
FontApptPrimary = CreateFont2(pctSchedule.hdc, "Arial", 8, True, False, False, False)
FontApptSecondary = CreateFont2(pctSchedule.hdc, "Arial", 8, False, False, False, False)
FontApptMinutes = CreateFont2(pctSchedule.hdc, "Arial", 8, False, False, False, False)
ReDim ColorProfilesDay(2, 3) '0=titletext, 1=titlebg, 2=border, 3=apptbg
'Past
ColorProfilesDay(0, 0) = &HC0C0C0
ColorProfilesDay(0, 1) = vbWhite
ColorProfilesDay(0, 2) = &H808080
ColorProfilesDay(0, 3) = vbWhite
'Today
ColorProfilesDay(1, 0) = vbBlack
ColorProfilesDay(1, 1) = vbGreen
ColorProfilesDay(1, 2) = &H8000&
ColorProfilesDay(1, 3) = vbWhite
'Future
ColorProfilesDay(2, 0) = vbBlack
ColorProfilesDay(2, 1) = &HFFFF&
ColorProfilesDay(2, 2) = vbBlack
ColorProfilesDay(2, 3) = vbWhite

ReDim ColorProfilesAppt(4, 3)  '0=text, 1=bg, 2=border, 3=postedcolor
'Normal
ColorProfilesAppt(0, 0) = vbBlack
ColorProfilesAppt(0, 1) = &HF0CAA6
ColorProfilesAppt(0, 2) = vbBlack
ColorProfilesAppt(0, 3) = &H808080
'Item being rescheduled
ColorProfilesAppt(1, 0) = &HC0C0C0
ColorProfilesAppt(1, 1) = vbYellow
ColorProfilesAppt(1, 2) = vbBlack
ColorProfilesAppt(1, 3) = ColorProfilesAppt(1, 0)
'Past
ColorProfilesAppt(2, 0) = vbBlack
ColorProfilesAppt(2, 1) = &HC0C0C0
ColorProfilesAppt(2, 2) = &H808080
ColorProfilesAppt(2, 3) = &H808080
'DidntHappen
ColorProfilesAppt(3, 0) = vbRed
ColorProfilesAppt(3, 1) = &H80FFFF
ColorProfilesAppt(3, 2) = &H808080
ColorProfilesAppt(3, 3) = ColorProfilesAppt(3, 0)
'Custom item
ColorProfilesAppt(4, 0) = vbBlack
ColorProfilesAppt(4, 1) = RGB(200, 255, 200)
ColorProfilesAppt(4, 2) = &H8000&
ColorProfilesAppt(4, 3) = ColorProfilesAppt(4, 0)

InitScheduleLayout

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CreateGDIObjects", Err
End Function

'EHT=Standard
Private Function ITab_InitializeAfterDBLoad() As Boolean
On Error GoTo ERR_HANDLER

ApptBeingRescheduled.ID = NullLong

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

If frmMain.CHOS_lstClients.ListCount = 0 Then
    ChangeScheduleMode sView
Else
    ChangeScheduleMode sCreate
End If
DrawSchedule

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr pctSchedule

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SetDefaultFocus", Err
End Sub

'EHT=Standard
Private Function ITab_SaveSettingsToDBBeforeClose() As Boolean
On Error GoTo ERR_HANDLER


Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SaveSettingsToDBBeforeClose", Err
End Function

'EHT=Standard
Private Function ITab_DestroyGDIObjects() As Boolean
On Error GoTo ERR_HANDLER

DeleteObject FontTitle
DeleteObject FontSubtitle
DeleteObject FontTimesOnSlot
DeleteObject FontTimesOffSlot
DeleteObject FontApptPrimary
DeleteObject FontApptSecondary
DeleteObject FontApptMinutes

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_DestroyGDIObjects", Err
End Function

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

pctSchedule.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
InitScheduleLayout
End Sub

'EHT=Standard
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

Select Case KeyCode
Case vbKeyHome
    ShowDate Date
Case vbKeyPageUp, vbKeyLeft, vbKeyUp
    ShowDate ViewStartDate - 7
Case vbKeyPageDown, vbKeyRight, vbKeyDown
    ShowDate ViewStartDate + 7
Case vbKeyAdd, 187
    If frmMain.CHOS_NumSlots > 0 Then
        frmMain.CHOS_NumSlots = frmMain.CHOS_NumSlots + 1
        frmMain.CHOS_NumSlots_Overridden = Not (frmMain.CHOS_NumSlots = frmMain.CHOS_NumSlotsBeforeOverride)
        frmMain.CHOS_UpdateTotal
        pctSchedule_MouseMove 0, Shift, lastMouseMoveX, lastMouseMoveY
    End If
Case vbKeySubtract, 189
    If frmMain.CHOS_NumSlots > 1 Then
        frmMain.CHOS_NumSlots = frmMain.CHOS_NumSlots - 1
        frmMain.CHOS_NumSlots_Overridden = Not (frmMain.CHOS_NumSlots = frmMain.CHOS_NumSlotsBeforeOverride)
        frmMain.CHOS_UpdateTotal
        pctSchedule_MouseMove 0, Shift, lastMouseMoveX, lastMouseMoveY
    End If
Case vbKeyBack
    frmMain.CHOS_CalculateTotal
    shpApptSelection.Visible = False
    lblApptSelection.Visible = False
Case vbKeyControl
    'In certain situations, the Ctrl key will change a Move to a Copy, but the MouseMove event won't know
    '  on initial Ctrl-press, so we handle that style change here. Of course, MouseDown will handle it properly
    '  so improper database changes won't happen regardless
    If LastShapeStyle = Style_MoveAndCtrlCopy Then pctSchedule_MouseMove 0, Shift, lastMouseMoveX, lastMouseMoveY
Case Else
    If IsLetterKey(KeyCode) Then
        'Send key to SRCH_txtSearch
        frmMain.ChangeCurTab vSearch, False
        tabSearch.ClearAll
        PutKeyCodeIntoTextbox tabSearch.txtSearch, KeyCode, True
    End If
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=Standard
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub     'Pass it to the parent form first, Exit if form cancelled the event
If KeyCode = vbKeyControl Then
    If LastShapeStyle = Style_CopyForcedWithCtrl Then pctSchedule_MouseMove 0, Shift, lastMouseMoveX, lastMouseMoveY
End If

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
Private Sub menApptEdit_Click()
On Error GoTo ERR_HANDLER

If Not menApptEdit.Enabled Then Exit Sub

Dim frm As frmApptEdit, aID&
aID = ActiveDBInstance.Appointments(ClickedApptIndex).ID
Set frm = New frmApptEdit
If frm.Form_Show(aID) Then         'This will mark changed if necessary
    DrawSchedule
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptEdit_Click", Err
End Sub

'EHT=Standard
Private Sub menApptReschedule_Click()
On Error GoTo ERR_HANDLER

If Not menApptReschedule.Enabled Then Exit Sub

ChangeScheduleMode sReschedule
ApptIDBeingRescheduled = ActiveDBInstance.Appointments(ClickedApptIndex).ID
DrawSchedule

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptReschedule_Click", Err
End Sub

'EHT=Standard
Private Sub menApptCancelDelete_Click()
On Error GoTo ERR_HANDLER

If Not menApptCancelDelete.Enabled Then Exit Sub

Dim a&, t$, c$, cindex&
With ActiveDBInstance.Appointments(ClickedApptIndex)
    'Format the time
    t$ = FormatApptTime(.ApptDate, .ApptActualTime)
    If .ClientID_Count = 0 Then
        If MsgBox("Are you sure you want to cancel '" & .Notes & "'?", _
                  vbQuestion Or vbYesNo) <> vbYes Then
            Exit Sub
        End If
    Else
        'Build list of clients
        For a = 0 To .ClientID_Count - 1
            cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(a))
            If Len(c$) > 0 Then c$ = c$ & vbCrLf & "             "
            c$ = c$ & FormatClientName(fSchedulePct, ActiveDBInstance.Clients(cindex).c)
        Next a
        If .ApptDate <= CLng(Date) Then
            If MsgBox("You are attempting to change an appointment on or prior to today. Please mark it as Didn't Happen instead of actually deleting it." & vbCrLf & vbCrLf & _
                      "Click Retry to ignore this warning and continue anyway.", _
                      vbCritical Or vbRetryCancel) <> vbRetry Then
                Exit Sub
            End If
        End If
        If MsgBox("Are you sure you want to cancel the following appointment?" & vbCrLf & _
                  "Time: " & t$ & vbCrLf & _
                  "Clients: " & c$, _
                  vbQuestion Or vbYesNo) <> vbYes Then
            Exit Sub
        End If
    End If
    For a = 0 To .ClientID_Count - 1
        cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(a))
        AddOpNote ActiveDBInstance.Clients(cindex).c.OpNotes, "Cancelled appt: " & t$
        ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
    Next a
    DB_SlotsClear ActiveDBInstance, .ApptDate, .ApptTimeSlot, .NumSlots
End With
'This must be outside of the with block for it to succeed
DB_RemoveAppointment ActiveDBInstance, ClickedApptIndex   'This will update the rest of the appt bitmap
frmMain.SetChangedFlagAndIndication
DrawSchedule
tabLogFile.WriteLine "Cancelled appt " & t$

ClickedApptIndex = -1   '...Since the apptindex doesn't exist anymore

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptCancelDelete_Click", Err
End Sub

'EHT=Standard
Private Sub menApptScheduleFromThis_Click()
On Error GoTo ERR_HANDLER

If Not menApptScheduleFromThis.Enabled Then Exit Sub

Dim a&, cindex&
frmMain.CHOS_lstClients.Clear
With ActiveDBInstance.Appointments(ClickedApptIndex)
    For a = 0 To .ClientID_Count - 1
        cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(a))
        frmMain.CHOS_Add2 .ClientIDs(a), cindex
    Next a
    frmMain.CHOS_CalculateTotal
    If frmMain.CHOS_NumSlots <> .NumSlots Then
        frmMain.CHOS_NumSlots = .NumSlots
        frmMain.CHOS_NumSlotsBeforeOverride = frmMain.CHOS_NumSlots
        frmMain.CHOS_NumSlots_Overridden = False
        frmMain.CHOS_UpdateTotal
        ChangeScheduleMode sCreate
    End If
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptScheduleFromThis_Click", Err
End Sub

'EHT=Standard
Private Sub menApptMarkReminderCall_Click()
On Error GoTo ERR_HANDLER

If Not menApptMarkReminderCall.Enabled Then Exit Sub

With ActiveDBInstance.Appointments(ClickedApptIndex)
    If Flag_IsSet(.Flags, ReminderCall) Then
        .Flags = Flag_Remove(.Flags, ReminderCall Or Called)    'Both
        tabLogFile.WriteLine "Marked NOT reminder call: " & DB_FormatApptClientList(ActiveDBInstance, ActiveDBInstance.Appointments(ClickedApptIndex))
    Else
        .Flags = .Flags Or ReminderCall
        tabLogFile.WriteLine "Marked reminder call: " & DB_FormatApptClientList(ActiveDBInstance, ActiveDBInstance.Appointments(ClickedApptIndex))
    End If
End With
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptMarkReminderCall_Click", Err
End Sub

'EHT=Standard
Private Sub menApptMarkCalled_Click()
On Error GoTo ERR_HANDLER

If Not menApptMarkCalled.Enabled Then Exit Sub

With ActiveDBInstance.Appointments(ClickedApptIndex)
    If Flag_IsSet(.Flags, Called) Then
        .Flags = Flag_Remove(.Flags, Called)
        tabLogFile.WriteLine "Marked NOT called: " & DB_FormatApptClientList(ActiveDBInstance, ActiveDBInstance.Appointments(ClickedApptIndex))
    Else
        .Flags = .Flags Or Called
        tabLogFile.WriteLine "Marked called: " & DB_FormatApptClientList(ActiveDBInstance, ActiveDBInstance.Appointments(ClickedApptIndex))
    End If
End With
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptMarkCalled_Click", Err
End Sub

'EHT=Standard
Private Sub menApptMarkDidntHappen_Click()
On Error GoTo ERR_HANDLER

If Not menApptMarkDidntHappen.Enabled Then Exit Sub

With ActiveDBInstance.Appointments(ClickedApptIndex)
    If Flag_IsSet(.Flags, DidntHappen) Then
        .Flags = Flag_Remove(.Flags, DidntHappen)
        tabLogFile.WriteLine "Marked DID happen: " & DB_FormatApptClientList(ActiveDBInstance, ActiveDBInstance.Appointments(ClickedApptIndex))
    Else
        .Flags = .Flags Or DidntHappen
        tabLogFile.WriteLine "Marked didn't happen: " & DB_FormatApptClientList(ActiveDBInstance, ActiveDBInstance.Appointments(ClickedApptIndex))
    End If
End With
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptMarkDidntHappen_Click", Err
End Sub

'EHT=Standard
Private Sub menApptCLItem_Click(Index As Integer)
On Error GoTo ERR_HANDLER

If Not menApptCLItem(Index).Enabled Then Exit Sub

Dim t$, cID&, cindex&
t$ = menApptCLItem(Index).Tag
If Len(t$) > 1 Then cID = Val(Mid$(t$, 2))
Select Case Mid$(t$, 1, 1)
Case "e"    'Edit
    Dim frme As frmClientEditPost
    Set frme = New frmClientEditPost
    If frme.Form_Show(cID, fEdit) Then  'This will mark changed if necessary
        frmMain.DayTotal_Update
        DrawSchedule
    End If
Case "p"    'Post
    Dim frmp As frmClientEditPost
    Set frmp = New frmClientEditPost
    If frmp.Form_Show(cID, fPost) Then      'This will mark changed if necessary
        frmMain.DayTotal_Update
        DrawSchedule
    End If
Case "i"    'Mark incomplete
    cindex = DB_FindClientIndex(ActiveDBInstance, cID)
    With ActiveDBInstance.Clients(cindex).c
        If Flag_IsSet(.Flags, PartiallyComplete) Then
            .Flags = Flag_Remove(.Flags, PartiallyComplete)
            AddOpNote .OpNotes, "Removed flag: INC"
            tabLogFile.WriteLine "Marked NOT INC: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
        Else
            .Flags = .Flags Or PartiallyComplete
            AddOpNote .OpNotes, "Incomplete"
            tabLogFile.WriteLine "Marked INC: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
        End If
        ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
    End With
    frmMain.SetChangedFlagAndIndication
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menApptCLItem_Click", Err
End Sub

'EHT=Standard
Private Sub menSlotMarkDefault_Click()
On Error GoTo ERR_HANDLER

If Not menSlotMarkDefault.Enabled Then Exit Sub

DB_SlotFill ActiveDBInstance, ClickedDate, ClickedTimeslot, Slot_DefaultAccordingToTemplate
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menSlotMarkDefault_Click", Err
End Sub

'EHT=Standard
Private Sub menSlotMarkAvailable_Click()
On Error GoTo ERR_HANDLER

If Not menSlotMarkAvailable.Enabled Then Exit Sub

DB_SlotFill ActiveDBInstance, ClickedDate, ClickedTimeslot, Slot_Available
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menSlotMarkAvailable_Click", Err
End Sub

'EHT=Standard
Private Sub menSlotMarkReserved_Click()
On Error GoTo ERR_HANDLER

If Not menSlotMarkReserved.Enabled Then Exit Sub

DB_SlotFill ActiveDBInstance, ClickedDate, ClickedTimeslot, Slot_Reserved
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menSlotMarkReserved_Click", Err
End Sub

'EHT=Standard
Private Sub menSlotMarkMealBreak_Click()
On Error GoTo ERR_HANDLER

If Not menSlotMarkMealBreak.Enabled Then Exit Sub

DB_SlotFill ActiveDBInstance, ClickedDate, ClickedTimeslot, Slot_MealBreak
DrawSchedule
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menSlotMarkMealBreak_Click", Err
End Sub

'EHT=Standard
Private Sub menSlotCreateNonClient_Click()
On Error GoTo ERR_HANDLER

If Not menSlotCreateNonClient.Enabled Then Exit Sub

If DB_SlotsIsAvail(ActiveDBInstance, ClickedDate, ClickedTimeslot, 1, -1) Then
    'Create appointment
    Dim a As Appointment, aindex&, frm As frmApptEdit
    Set frm = New frmApptEdit
    With a
        .ID = DB_GetNewAppointmentID(ActiveDBInstance)
        .ApptDate = ClickedDate
        .ApptTimeSlot = ClickedTimeslot
        .ApptActualTime = Appointment_FirstSlotTime + (.ApptTimeSlot * Appointment_SlotLength)
        .NumSlots = 1
        .Notes = ""
    End With
    'Add it to the schedule
    aindex = DB_AddAppointment(ActiveDBInstance, a)
    DB_SlotFill ActiveDBInstance, a.ApptDate, a.ApptTimeSlot, aindex
    frmMain.SetChangedFlagAndIndication
    DrawSchedule
    If frm.Form_Show(a.ID, True) Then        'This will mark changed if necessary
        DrawSchedule
    Else
        'User canceled the initial edit, so remove the new item from the schedule
        DB_SlotsClear ActiveDBInstance, a.ApptDate, a.ApptTimeSlot, a.NumSlots
        'This must be outside of the with block for it to succeed
        DB_RemoveAppointment ActiveDBInstance, aindex 'This will update the rest of the appt bitmap
        frmMain.SetChangedFlagAndIndication
        DrawSchedule
    End If
Else
    Err.Raise 1, , "User attempted to create a custom item in a slot which is unavailable."
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menSlotCreateNonClient_Click", Err
End Sub

'EHT=Standard
Private Sub pctSchedule_DblClick()
On Error GoTo ERR_HANDLER

If DoubleClickAllowed Then
    If ClickedApptIndex >= 0 Then
        If ActiveDBInstance.Appointments(ClickedApptIndex).ClientID_Count > 0 Then
            Dim frme As frmClientEditPost
            Set frme = New frmClientEditPost
            If frme.Form_Show(ActiveDBInstance.Appointments(ClickedApptIndex).ClientIDs(0), fEdit) Then   'This will mark changed if necessary
                frmMain.DayTotal_Update
                DrawSchedule
            End If
        Else
            Dim frm As frmApptEdit, aID&
            aID = ActiveDBInstance.Appointments(ClickedApptIndex).ID
            Set frm = New frmApptEdit
            If frm.Form_Show(aID) Then         'This will mark changed if necessary
                DrawSchedule
            End If
        End If
    ElseIf ClickedTimeslot >= 0 Then
        If ActiveDBInstance.IsWriteable Then menSlotCreateNonClient_Click
    ElseIf ClickedDate >= 0 Then
        If ActiveDBInstance.IsWriteable Then
            Dim i%, t$
            i = ClickedDate - ActiveDBInstance.ApptBitmap_StartDate
            t$ = ActiveDBInstance.Subtitles(i)
            t$ = InputBox("Edit the subtitle for " & FormatDateForDayTitle$(ClickedDate) & vbCrLf & vbCrLf & "Enter a hyphen (-) to remove the subtitle.", , t$)
            If Len(t$) > 0 Then
                If t$ = "-" Then
                    ActiveDBInstance.Subtitles(i) = ""
                Else
                    ActiveDBInstance.Subtitles(i) = t$
                End If
                DrawSchedule
                frmMain.SetChangedFlagAndIndication
            End If
        End If
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "pctSchedule_DblClick", Err
End Sub

'EHT=Standard
Private Sub pctSchedule_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER

pctSchedule_MouseMove Button, Shift, X, Y

Dim aindex&, b&, cindex&, t$, cli&, comp As Boolean
Dim a As Appointment, defmi As Menu
Dim abr&, moveexistingappt As ScheduleShapeStyle

'Don't clear ClickedApptIndex at this point, but make sure that every possible code path below uses/sets it properly

MouseMoveCalc X, Y, ClickedDayIndex, ClickedDate, ClickedTimeslot
DoubleClickAllowed = False
If ClickedTimeslot >= 0 Then
    Select Case ScheduleMode
    Case sView
        ClickedApptIndex = ActiveDBInstance.ApptBitmap(ClickedDate - ActiveDBInstance.ApptBitmap_StartDate, ClickedTimeslot)
        If ClickedApptIndex < 0 Then
            'Over a blank timeslot
            If Button = vbLeftButton Then
                'Let the double-click through so the user can create a new non-client appointment
                DoubleClickAllowed = True
            ElseIf Button = vbRightButton Then
                'Show Slot menu
                menSlot_Title.Caption = "=== " & DB_GetTimeSlotTime(ClickedTimeslot) & " ==="
                menSlotMarkDefault.Checked = (ClickedApptIndex = Slot_DefaultAccordingToTemplate)
                menSlotMarkDefault.Enabled = (ActiveDBInstance.IsWriteable)
                menSlotMarkAvailable.Checked = (ClickedApptIndex = Slot_Available)
                menSlotMarkAvailable.Enabled = (ActiveDBInstance.IsWriteable)
                menSlotMarkReserved.Checked = (ClickedApptIndex = Slot_Reserved)
                menSlotMarkReserved.Enabled = (ActiveDBInstance.IsWriteable)
                menSlotMarkMealBreak.Checked = (ClickedApptIndex = Slot_MealBreak)
                menSlotMarkMealBreak.Enabled = (ActiveDBInstance.IsWriteable)
                menSlotCreateNonClient.Enabled = (ActiveDBInstance.IsWriteable)
                PopupMenu menSlot   'No 'With' blocks!!!
            End If
        Else 'ClickedApptIndex >= 0
            'Over an existing appointment
            If Button = vbLeftButton Then
                'Let the double-click through so the user can edit the appointment
                DoubleClickAllowed = True
            ElseIf Button = vbRightButton Then
                'Show Appt menu
                ClickedTimeslot = ClickedTimeslot
                With ActiveDBInstance.Appointments(ClickedApptIndex)
                    menAppt_Title.Caption = "======= " & Format$(.ApptActualTime, "h:mm AM/PM") & " ======="
                    menApptReschedule.Enabled = (ActiveDBInstance.IsWriteable)
                    menApptCancelDelete.Caption = IIf(.ClientID_Count = 0, "&Delete", "&Cancel")
                    menApptCancelDelete.Enabled = (ActiveDBInstance.IsWriteable)
                    menApptScheduleFromThis.Enabled = (.ClientID_Count > 0) And (ActiveDBInstance.IsWriteable)
                    menApptMarkReminderCall.Checked = Flag_IsSet(.Flags, ReminderCall)
                    menApptMarkReminderCall.Enabled = (.ClientID_Count > 0) And (ActiveDBInstance.IsWriteable)
                    menApptMarkCalled.Checked = Flag_IsSet(.Flags, Called)
                    menApptMarkCalled.Enabled = (.ClientID_Count > 0) And menApptMarkReminderCall.Checked And (ActiveDBInstance.IsWriteable)
                    menApptMarkDidntHappen.Checked = Flag_IsSet(.Flags, DidntHappen)
                    menApptMarkDidntHappen.Enabled = (.ClientID_Count > 0) And (ActiveDBInstance.IsWriteable)
                    cli = 0
                    For b = 0 To .ClientID_Count - 1
                        cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(b))
                        comp = Flag_IsSet(ActiveDBInstance.Clients(cindex).c.Flags, CompletedReturn) Or Flag_IsSet(ActiveDBInstance.Clients(cindex).c.Flags, NoNeedToFile)

                        cli = cli + 1
                        Load menApptCLItem(cli)
                        menApptCLItem(cli).Caption = "-"
                        menApptCLItem(cli).Visible = True

                        cli = cli + 1
                        Load menApptCLItem(cli)
                        menApptCLItem(cli).Caption = "== " & Replace(FormatClientName(fSchedulePct, ActiveDBInstance.Clients(cindex).c), "&", "&&") & " =="
                        menApptCLItem(cli).Enabled = False
                        menApptCLItem(cli).Visible = True

                        cli = cli + 1
                        Load menApptCLItem(cli)
                        menApptCLItem(cli).Caption = "Edit Client..."
                        menApptCLItem(cli).Visible = True
                        menApptCLItem(cli).Tag = "e" & .ClientIDs(b)
                        If b = 0 Then Set defmi = menApptCLItem(cli)

                        cli = cli + 1
                        Load menApptCLItem(cli)
                        menApptCLItem(cli).Caption = "Post Client..."
                        menApptCLItem(cli).Enabled = Not comp
                        menApptCLItem(cli).Visible = True
                        menApptCLItem(cli).Tag = "p" & .ClientIDs(b)

                        cli = cli + 1
                        Load menApptCLItem(cli)
                        menApptCLItem(cli).Caption = "Incomplete"
                        menApptCLItem(cli).Checked = Flag_IsSet(ActiveDBInstance.Clients(cindex).c.Flags, PartiallyComplete)
                        menApptCLItem(cli).Enabled = (Not comp) And (ActiveDBInstance.IsWriteable)
                        menApptCLItem(cli).Visible = True
                        menApptCLItem(cli).Tag = "i" & .ClientIDs(b)
                    Next b
                End With
                'Must be outside the With block
                If defmi Is Nothing Then Set defmi = menApptEdit
                PopupMenu menAppt, , , , defmi        'No With blocks!!!
                cli = menApptCLItem.UBound  'Keep this separate from For Loop
                For b = 1 To cli
                    Unload menApptCLItem(b)
                Next b
            End If
        End If
    Case sCreate
        ClickedApptIndex = -1
        If Not ActiveDBInstance.IsWriteable Then Exit Sub
        If DB_SlotsIsAvail(ActiveDBInstance, ClickedDate, ClickedTimeslot, frmMain.CHOS_NumSlots, -1) Then
            If ClickedDate = Date Then
                If MsgBox("Adding appointments to current day is not allowed! Continue anyway?", vbCritical Or vbYesNo Or vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            End If

            'Create appointment
            Dim n As Boolean
            With a
                .ID = DB_GetNewAppointmentID(ActiveDBInstance)
                .ApptDate = ClickedDate
                .ApptTimeSlot = ClickedTimeslot
                .ApptActualTime = Appointment_FirstSlotTime + (.ApptTimeSlot * Appointment_SlotLength)
                .NumSlots = frmMain.CHOS_NumSlots
                t$ = FormatApptTime$(.ApptDate, .ApptActualTime)
                .ClientID_Count = frmMain.CHOS_lstClients.ListCount
                If .ClientID_Count = 0 Then
                    Erase .ClientIDs
                Else
                    ReDim .ClientIDs(.ClientID_Count - 1)
                    For b = 0 To .ClientID_Count - 1
                        .ClientIDs(b) = frmMain.CHOS_lstClients.ItemData(b)
                        cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(b))
                        If Flag_IsSet(ActiveDBInstance.Clients(cindex).c.Flags, NewClient) Then n = True
                        AddOpNote ActiveDBInstance.Clients(cindex).c.OpNotes, "Scheduled appt: " & t$
                        ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
                    Next b
                    If n Then
                        .Flags = ReminderCall
                    ElseIf ((ClickedDate - CLng(Date)) > DB_GetSetting(ActiveDBInstance, "Reminder call if appt scheduled more than")) Then
                        .Flags = ReminderCall
                    End If
                End If
                aindex = DB_AddAppointment(ActiveDBInstance, a)

                DB_SlotsFill ActiveDBInstance, .ApptDate, .ApptTimeSlot, .NumSlots, aindex
                frmMain.DayTotal_Update
                frmMain.SetChangedFlagAndIndication
                ChangeScheduleMode sView
                frmMain.CHOS_Clear
                tabSchedule.ChangeScheduleMode sView
                DrawSchedule
                tabLogFile.WriteLine "Created " & DB_FormatApptClientList(ActiveDBInstance, a) & ": " & t$ & ", " & FormatNumApptSlots(.NumSlots)
            End With
        End If
    Case sReschedule
        'At this point, ClickedApptIndex is the index of the original appointment we're trying to reschedule from
        If Not ActiveDBInstance.IsWriteable Then Exit Sub
        a = ActiveDBInstance.Appointments(ClickedApptIndex)
        moveexistingappt = IsMoveOrCopy(a, ClickedDate, Shift)
        If moveexistingappt = Style_Move Or moveexistingappt = Style_MoveAndCtrlCopy Then
            'Since we're moving the appointment, we can 'overlap' the old position
            abr = ApptIDBeingRescheduled
        Else
            'But if copying, overlap is prohibited
            abr = -1
        End If
        If DB_SlotsIsAvail(ActiveDBInstance, ClickedDate, ClickedTimeslot, ActiveDBInstance.Appointments(ClickedApptIndex).NumSlots, abr) Then
            Dim oldat$

            'Create new appt structure
            With a
                oldat$ = FormatApptTime$(.ApptDate, .ApptActualTime)
                If moveexistingappt = Style_Copy Or moveexistingappt = Style_CopyForcedWithCtrl Then .ID = DB_GetNewAppointmentID(ActiveDBInstance)
                .ApptDate = ClickedDate
                .ApptTimeSlot = ClickedTimeslot
                .ApptActualTime = Appointment_FirstSlotTime + (.ApptTimeSlot * Appointment_SlotLength)
                .Flags = Flag_Remove(.Flags, DidntHappen)
                t$ = FormatApptTime$(.ApptDate, .ApptActualTime)
                For b = 0 To .ClientID_Count - 1
                    cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(b))
                    AddOpNote ActiveDBInstance.Clients(cindex).c.OpNotes, "Resch appt to: " & t$
                    ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
                Next b
                If .ClientID_Count > 0 Then
                    If ((ClickedDate - CLng(Date)) > DB_GetSetting(ActiveDBInstance, "Reminder call if appt scheduled more than")) Then a.Flags = ReminderCall
                End If
            End With

            'Handle old appt
            With ActiveDBInstance.Appointments(ClickedApptIndex)
                If moveexistingappt = Style_Copy Or moveexistingappt = Style_CopyForcedWithCtrl Then
                    'Mark that the original appointment didn't happen
                    If .ClientID_Count > 0 Then .Flags = .Flags Or DidntHappen
                Else
                    'Remove the old appointment
                    DB_SlotsClear ActiveDBInstance, .ApptDate, .ApptTimeSlot, .NumSlots
                End If
            End With

            'Put new appt into database
            If moveexistingappt = Style_Copy Or moveexistingappt = Style_CopyForcedWithCtrl Then
                'Create new appointment
                aindex = DB_AddAppointment(ActiveDBInstance, a)
                DB_SlotsFill ActiveDBInstance, a.ApptDate, a.ApptTimeSlot, a.NumSlots, aindex
            Else
                'Overwrite old appointment with new information
                ActiveDBInstance.Appointments(ClickedApptIndex) = a
                DB_SlotsFill ActiveDBInstance, a.ApptDate, a.ApptTimeSlot, a.NumSlots, ClickedApptIndex
            End If

            'Finish
            frmMain.SetChangedFlagAndIndication
            ChangeScheduleMode sView
            DrawSchedule
            tabLogFile.WriteLine "Rescheduled " & DB_FormatApptClientList(ActiveDBInstance, a) & ": " & oldat$ & " to " & FormatApptTime$(a.ApptDate, a.ApptActualTime)
        End If
    End Select
Else 'ClickedTimeslot < 0
    ClickedApptIndex = -1
    If Button = vbLeftButton Then
        If ScheduleMode = sView Then
            'Double-clicked on a non-timeslot area within a day rectangle (probably the title)
            'Let the double-click through so the user can change the subtitle text
            DoubleClickAllowed = True
        End If
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "pctSchedule_MouseDown", Err
End Sub

'EHT=Standard
Private Sub pctSchedule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER

Dim moDayIndex&, moDate&, moTimeSlot&
Dim a As Appointment, ai&, ns&, abr&, moveexistingappt As ScheduleShapeStyle

'Make sure this happens first...
'Several places in this module must call the MouseMove event later, without the mouse actually moving
'  so we save the coordinates here for that purpose
lastMouseMoveX = X
lastMouseMoveY = Y

If tmrFlashAppt.Enabled Then
    If MouseNullZone_Moved() Then
        StopFlashAppt
    Else
        Exit Sub
    End If
End If

MouseMoveCalc X, Y, moDayIndex, moDate, moTimeSlot
If moTimeSlot < 0 Then
    'Not over a valid time slot
    HideShape
Else
    ai = ActiveDBInstance.ApptBitmap(moDate - ActiveDBInstance.ApptBitmap_StartDate, moTimeSlot)
    'frmMain.Caption = ai
    Select Case ScheduleMode
    Case sView
        If ai < 0 Then
            HideShape
            Exit Sub 'Don't hilight an empty slot in view mode
        Else
            moTimeSlot = ActiveDBInstance.Appointments(ai).ApptTimeSlot
            ns = ActiveDBInstance.Appointments(ai).NumSlots
            MoveShapeAndSetStyle moDayIndex, moTimeSlot, ns, Style_Normal
        End If
    Case sCreate
        If Not ActiveDBInstance.IsWriteable Then
            HideShape
            Exit Sub
        Else
            If DB_SlotsIsAvail(ActiveDBInstance, moDate, moTimeSlot, frmMain.CHOS_NumSlots, -1) Then
                ns = frmMain.CHOS_NumSlots
                MoveShapeAndSetStyle moDayIndex, moTimeSlot, ns, Style_New
            Else
                HideShape
                Exit Sub 'Can't do anything with existing appointments in create or reschedule modes
            End If
        End If
    Case sReschedule
        If Not ActiveDBInstance.IsWriteable Then
            HideShape
            Exit Sub
        Else
            a = ActiveDBInstance.Appointments(ClickedApptIndex)
            moveexistingappt = IsMoveOrCopy(a, moDate, Shift)
            If moveexistingappt = Style_Move Or moveexistingappt = Style_MoveAndCtrlCopy Then
                'Since we're moving the appointment, we can 'overlap' the old position
                abr = ApptIDBeingRescheduled
            Else
                'But if copying, overlap is prohibited
                abr = -1
            End If
            If DB_SlotsIsAvail(ActiveDBInstance, moDate, moTimeSlot, a.NumSlots, abr) Then
                ns = a.NumSlots
                If moveexistingappt = Style_MoveAndCtrlCopy Then
                    'This is a special situation, because if the user is hilighting a slot that
                    '  overlaps ApptIDBeingRescheduled, Move is fine but Copy is not. So we must
                    '  determine this ahead of time. If Copy would not be allowed in this position
                    '  if the user hits Ctrl, then it's better to change Style_MoveAndCtrlCopy to
                    '  Style_Move to prevent the user from even hitting Ctrl in the first place.
                    If DB_SlotsIsAvail(ActiveDBInstance, moDate, moTimeSlot, a.NumSlots, -1) Then
                        'Slots are wide open, so we can continue with Style_MoveAndCtrlCopy
                        MoveShapeAndSetStyle moDayIndex, moTimeSlot, ns, moveexistingappt
                    Else
                        'This position is actually overlapping ApptIDBeingRescheduled, so change the
                        '  mode to Style_Move to prevent the Ctrl toggle from happening in Form_KeyDown
                        MoveShapeAndSetStyle moDayIndex, moTimeSlot, ns, Style_Move
                    End If
                Else
                    MoveShapeAndSetStyle moDayIndex, moTimeSlot, ns, moveexistingappt
                End If
            Else
                HideShape
                Exit Sub 'Can't do anything with existing appointments in create or reschedule modes
            End If
        End If
    End Select
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "pctSchedule_MouseMove", Err
End Sub

'EHT=ResumeNext
Private Sub tmrFlashAppt_Timer()
On Error Resume Next

If Timer < FlashStopTime Then
    shpApptSelection.Visible = Not shpApptSelection.Visible
Else
    StopFlashAppt
End If
End Sub

'EHT=Standard
Sub InitScheduleLayout()
On Error GoTo ERR_HANDLER

Dim r&, c&, i&

Select Case pctSchedule.Height
Case Is < 378
    DayApptSlotHeight = 10
Case Is < 406
    DayApptSlotHeight = 11
Case Is < 434
    DayApptSlotHeight = 12
Case Is < 462
    DayApptSlotHeight = 13
Case Is < 491
    DayApptSlotHeight = 14
Case Is < 518
    DayApptSlotHeight = 15
Case Is < 547
    DayApptSlotHeight = 16
Case Else
    DayApptSlotHeight = 17
End Select
DayHeight = DayFirstSlotOffsetY + (DayApptSlotHeight * Appointment_NumSlots) + 1 + DayMarginBottom + 1

i = 0
                        #If LRTB Then
For r = 0 To 1
    For c = 0 To 2
                        #Else
For c = 0 To 2
    For r = 0 To 1
                        #End If
        With ScheduleDayPositions(i)
            .Left = MarginLeft + (c * (DayWidth + DaySpacingX))
            .Right = .Left + DayWidth - 1
            .Top = MarginTop + (r * (DayHeight + DaySpacingY))
            .Bottom = .Top + DayHeight - 1
        End With
        i = i + 1
    Next
Next
With ScheduleDayPositions(6)
    .Left = MarginLeft + (3 * (DayWidth + DaySpacingX))
    .Right = .Left + DayWidth - 1
    .Top = MarginTop + (1 * (DayHeight + DaySpacingY))
    .Bottom = .Top + DayHeight - 1
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "InitScheduleLayout", Err
End Sub

'EHT=Standard
Sub DrawSchedule()
On Error GoTo ERR_HANDLER

If Not ActiveDBInstance.Loaded Then Exit Sub

Dim a&, todaysdate&

ExtractScheduleTemplate

pctSchedule.Cls
pctScheduleHdc = pctSchedule.hdc
TodaysDayIndex = CalcTodaysDayIndex(todaysdate)
For a = 0 To 6
    DrawScheduleDay ViewStartDate + a, ScheduleDayPositions(a).Left, ScheduleDayPositions(a).Top, todaysdate
Next a

MoveRedArrow Time

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DrawSchedule", Err
End Sub

'EHT=Standard
Sub DrawScheduleDay(cd As Long, cx&, cy&, todaysdate As Long)
On Error GoTo ERR_HANDLER

Dim a&, b&, c&, cindex&, tx&, ty&, t$, ts&
Dim scheduletemplaterange&
Dim CurCPDay&, CurCPAppt&
Dim r As RECT

If Not DB_DayWithinBitmapRange(ActiveDBInstance, cd) Then Exit Sub

Select Case cd
Case Is < todaysdate
    CurCPDay = 0
Case todaysdate
    CurCPDay = 1
Case Is > todaysdate
    CurCPDay = 2
End Select

If cd < DB_GetSetting(ActiveDBInstance, "Schedule Template B starting date") Then
    scheduletemplaterange = 1
ElseIf cd >= DB_GetSetting(ActiveDBInstance, "Schedule Template C starting date") Then
    scheduletemplaterange = 3
Else
    scheduletemplaterange = 2
End If

'Title background
pctSchedule.Line (cx, cy)-(cx + DayWidth - 1, cy + DayTitleHeight - 1), ColorProfilesDay(CurCPDay, 1), BF

'Background
pctSchedule.Line (cx, cy + DayTitleHeight - 1)-(cx + DayWidth - 1, cy + DayHeight - 1), ColorProfilesDay(CurCPDay, 3), BF

'Border
If CurCPDay = 1 Then pctSchedule.Line (cx - 1, cy - 1)-(cx + DayWidth, cy + DayHeight), ColorProfilesDay(CurCPDay, 2), B
pctSchedule.Line (cx, cy)-(cx + DayWidth - 1, cy + DayHeight - 1), ColorProfilesDay(CurCPDay, 2), B

'Title separator
pctSchedule.Line (cx + 1, cy + DayTitleHeight - 1)-(cx + DayWidth - 1, cy + DayTitleHeight - 1), ColorProfilesDay(CurCPDay, 2)

'Draw title
SetTextColor pctScheduleHdc, ColorProfilesDay(CurCPDay, 0)
SelectObject pctScheduleHdc, FontTitle
SetTextAlign pctScheduleHdc, TA_CENTER
t$ = FormatDateForDayTitle$(cd)
TextOut pctScheduleHdc, cx + (DayWidth / 2), cy + 1, t$, Len(t$)

'Draw subtitle
SetTextColor pctScheduleHdc, ColorProfilesDay(CurCPDay, 0)
SelectObject pctScheduleHdc, FontSubtitle
SetTextAlign pctScheduleHdc, TA_CENTER
t$ = ActiveDBInstance.Subtitles(cd - ActiveDBInstance.ApptBitmap_StartDate)
TextOut pctScheduleHdc, cx + (DayWidth / 2), cy + 23, t$, Len(t$)

'Draw time slots
SelectObject pctScheduleHdc, FontTimesOnSlot
For ts = 0 To Appointment_NumSlotsUB
    b = ActiveDBInstance.ApptBitmap(cd - ActiveDBInstance.ApptBitmap_StartDate, ts)
    ty = cy + DayFirstSlotOffsetY + (ts * DayApptSlotHeight)
    If b = Slot_DefaultAccordingToTemplate Then
        'Lookup the schedule template in the settings for that day and slot
        b = ScheduleTemplate(scheduletemplaterange, cd - ViewStartDate, ts)
    End If
    Select Case b
    Case Slot_Reserved
        pctSchedule.FillStyle = FillStyleConstants.vbDownwardDiagonal
        pctSchedule.DrawStyle = DrawStyleConstants.vbInvisible
        pctSchedule.Line (cx + DayApptsOffsetX, ty + 1)-(cx + DayWidth - DayMarginRight - 2, ty + DayApptSlotHeight + 1), ColorSlotReserved, B
        pctSchedule.FillStyle = FillStyleConstants.vbFSTransparent
        pctSchedule.DrawStyle = DrawStyleConstants.vbSolid
    Case Slot_MealBreak
        pctSchedule.Line (cx + DayMarginLeft + 1, ty + 1)-(cx + DayWidth - DayMarginRight - 2, ty + DayApptSlotHeight), ColorSlotMealBreak, BF
        SetTextAlign pctScheduleHdc, TA_CENTER
        SetTextColor pctScheduleHdc, vbBlack
        t$ = "MEAL BREAK"
        TextOut pctScheduleHdc, cx + DayApptsOffsetX + ((DayWidth - DayApptsOffsetX - DayMarginRight) / 2), ty + 1, t$, Len(t$)
    End Select
    If b < 0 Then
        'Normal slot, time on left side, if empty
        SetTextAlign pctScheduleHdc, TA_RIGHT
        SetTextColor pctScheduleHdc, ColorTimeText_Empty
        t$ = Format$(CDate((Appointment_FirstSlotTime + (ts * Appointment_SlotLength))), "h:mma/p")
        TextOut pctScheduleHdc, cx + DayTimesOffsetX, ty, t$, Len(t$)
    End If
Next ts

'Draw appointments
SetTextAlign pctScheduleHdc, TA_LEFT
For a = 0 To ActiveDBInstance.Appointments_Count - 1
    With ActiveDBInstance.Appointments(a)
        If .ApptDate = cd Then
            'Choose color profile
            If (ScheduleMode = sReschedule) And (.ID = ApptIDBeingRescheduled) Then
                CurCPAppt = 1
            ElseIf .ClientID_Count = 0 Then
                CurCPAppt = 4
            ElseIf Flag_IsSet(.Flags, DidntHappen) Then
                CurCPAppt = 3
            ElseIf CurCPDay = 0 Then
                CurCPAppt = 2
            Else
                CurCPAppt = 0
            End If

            'Calculate position on pctSchedule
            tx = cx + DayApptsOffsetX
            ty = cy + DayFirstSlotOffsetY + (.ApptTimeSlot * DayApptSlotHeight)

            'Draw appointment time (left)
            SetTextAlign pctScheduleHdc, TA_RIGHT
            If Appointment_FirstSlotTime + (.ApptTimeSlot * Appointment_SlotLength) = .ApptActualTime Then
                'Time is on-slot
                SelectObject pctScheduleHdc, FontTimesOnSlot
                SetTextColor pctScheduleHdc, ColorTimeText_OnSlot
            Else
                'Time is off-slot
                SelectObject pctScheduleHdc, FontTimesOffSlot
                SetTextColor pctScheduleHdc, ColorTimeText_OffSlot
            End If
            t$ = Format$(.ApptActualTime, "h:mma/p")
            TextOut pctScheduleHdc, cx + DayTimesOffsetX, ty, t$, Len(t$)

            If .ClientID_Count > 0 Then
                'Draw appointment rectangle
                pctSchedule.Line (tx, ty)-(cx + DayWidth - DayMarginRight - 2, ty + (.NumSlots * DayApptSlotHeight)), ColorProfilesAppt(CurCPAppt, 1), BF
                pctSchedule.Line (tx, ty)-(cx + DayWidth - DayMarginRight - 2, ty + (.NumSlots * DayApptSlotHeight)), ColorProfilesAppt(CurCPAppt, 2), B

                'Draw clients
                c = .ClientID_Count - 1
                If c > (.NumSlots - 1) Then
                    c = .NumSlots - 1
                End If
                For b = 0 To c
                    cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(b))

                    'Name
                    SetTextAlign pctScheduleHdc, TA_LEFT
                    If Flag_IsSet(ActiveDBInstance.Clients(cindex).c.Flags, CompletedReturn) Then
                        SetTextColor pctScheduleHdc, ColorProfilesAppt(CurCPAppt, 3)
                    Else
                        SetTextColor pctScheduleHdc, ColorProfilesAppt(CurCPAppt, 0)
                    End If
                    If b = 0 Then
                        SelectObject pctScheduleHdc, FontApptPrimary
                        t$ = FormatClientName(fSchedulePct, ActiveDBInstance.Clients(cindex).c)
                    Else
                        SelectObject pctScheduleHdc, FontApptSecondary
                        t$ = "+ " & FormatClientName(fSchedulePct, ActiveDBInstance.Clients(cindex).c)
                    End If
                    If (b = c) And (.ClientID_Count > .NumSlots) Then
                        t$ = t$ & " + ..............."
                    End If
                    r.Left = tx + 3
                    r.Top = ty + (b * DayApptSlotHeight) + 1
                    r.Right = cx + DayWidth - DayMarginRight - 2 - 18
                    r.Bottom = ty + ((b + 1) * DayApptSlotHeight) + 1
                    DrawText pctScheduleHdc, t$, Len(t$), r, DT_NOPREFIX Or DT_LEFT Or DT_WORD_ELLIPSIS

                    'Last year's minutes
                    SetTextAlign pctScheduleHdc, TA_RIGHT
                    SelectObject pctScheduleHdc, FontApptMinutes
                    'SetTextColor pctScheduleHdc, ColorProfilesAppt(CurCPAppt, 3)
                    t$ = DB_FormatMinutesForSchedule(ActiveDBInstance, cindex, b = 0)
                    If (b = c) And (.ClientID_Count > .NumSlots) Then
                        t$ = t$ & " +..."
                    End If
                    TextOut pctScheduleHdc, cx + DayWidth - DayMarginRight - 4, ty + (b * DayApptSlotHeight) + 1, t$, Len(t$)
                Next b
            Else
                'Custom item, no clients

                'Draw appointment rectangle
                pctSchedule.Line (tx, ty)-(cx + DayWidth - DayMarginRight - 2, ty + (.NumSlots * DayApptSlotHeight)), ColorProfilesAppt(CurCPAppt, 1), BF
                pctSchedule.Line (tx, ty)-(cx + DayWidth - DayMarginRight - 2, ty + (.NumSlots * DayApptSlotHeight)), ColorProfilesAppt(CurCPAppt, 2), B

                'Draw appointment notes
                SetTextAlign pctScheduleHdc, TA_LEFT
                SetTextColor pctScheduleHdc, ColorProfilesAppt(CurCPAppt, 0)
                SelectObject pctScheduleHdc, FontTimesOnSlot
                t$ = .Notes
                r.Left = tx + 2
                r.Top = ty + 1
                r.Right = cx + DayWidth - DayMarginRight - 2 - 1
                r.Bottom = ty + (.NumSlots * DayApptSlotHeight) - 1
                DrawText pctScheduleHdc, t$, Len(t$), r, DT_NOPREFIX Or DT_CENTER Or DT_WORDBREAK
            End If

            'Draw reminder call flag
            If Flag_IsSet(.Flags, ReminderCall) Then
                pctSchedule.PaintPicture imgReminderCall(Flag_IsSet(.Flags, Called) + 1).Picture, cx + DayTimesOffsetX + 1, ty + 3
            End If
        End If
    End With
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DrawScheduleDay", Err
End Sub

'EHT=Standard
Sub ExtractScheduleTemplate()
On Error GoTo ERR_HANDLER

Dim r&, wd&, ts&, s$
For r = 1 To 3
    For wd = 0 To 6
        s$ = UCase$(DB_GetSetting(ActiveDBInstance, "Schedule Template " & Chr(64 + r) & (wd + 1) & " (" & WeekdayName(wd + 1, False, vbMonday) & ")"))
        For ts = 1 To Appointment_NumSlots
            Select Case Mid$(s$, ts, 1)
            Case "R"
                ScheduleTemplate(r, wd, ts - 1) = Slot_Reserved
            Case "M"
                ScheduleTemplate(r, wd, ts - 1) = Slot_MealBreak
            Case Else  ' "A" or anything incorrect
                ScheduleTemplate(r, wd, ts - 1) = Slot_Available
            End Select
        Next ts
    Next wd
Next r

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ExtractScheduleTemplate", Err
End Sub

'EHT=Standard
Private Function IsMoveOrCopy(OrigAppt As Appointment, NewDate As Long, ShiftState As Integer) As ScheduleShapeStyle
On Error GoTo ERR_HANDLER

'Returns a style code, from the SetShapeStyle function

Dim todaydate As Long
todaydate = CLng(Date)
If OrigAppt.ClientID_Count = 0 Then
    'Non-client appointment
    If (ShiftState And vbCtrlMask) = vbCtrlMask Then
        'If holding Ctrl
        IsMoveOrCopy = Style_CopyForcedWithCtrl
    Else
        'Otherwise
        IsMoveOrCopy = Style_MoveAndCtrlCopy
    End If
Else
    'Client appointment
    If OrigAppt.ApptDate < todaydate Then        'Past
        'To anything
        IsMoveOrCopy = Style_Copy
    ElseIf OrigAppt.ApptDate = todaydate Then    'Present
        If NewDate = todaydate Then
            'To present
            IsMoveOrCopy = Style_Move
        Else
            'To anything else
            IsMoveOrCopy = Style_Copy
        End If
    Else                                        'Future
        'To anything
        IsMoveOrCopy = Style_Move
    End If
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "IsMoveOrCopy", Err
End Function

'EHT=Standard
Function CalcTodaysDayIndex(ByRef todaysdate As Long) As Long
On Error GoTo ERR_HANDLER

todaysdate = CLng(Date)
CalcTodaysDayIndex = todaysdate - ViewStartDate
If (CalcTodaysDayIndex < 0) Or (CalcTodaysDayIndex > 6) Then
    'Today not visible
    CalcTodaysDayIndex = -1
Else
    'Today not within DB's bitmap range
    If Not DB_DayWithinBitmapRange(ActiveDBInstance, todaysdate) Then
        CalcTodaysDayIndex = -1
    End If
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CalcTodaysDayIndex", Err
End Function

'EHT=Standard
Public Sub ChangeScheduleMode(sm As enumScheduleMode, Optional appt)
On Error GoTo ERR_HANDLER

ScheduleMode = sm
ApptIDBeingRescheduled = -1
shpApptSelection.Visible = False
lblApptSelection.Visible = False

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ChangeScheduleMode", Err
End Sub

'EHT=Standard
Sub MouseMoveCalc(ByVal X As Long, ByVal Y As Long, ByRef moDayIndex&, ByRef moDate&, ByRef moTimeSlot&)
On Error GoTo ERR_HANDLER

'If cursor is over a day on the schedule, returns Index and Date of it; otherwise -1
'If cursor is over a time slot, alse returns time slot index; otherwise -1

Dim cx&, cy&                'Cursor position relative to the first day
Dim colindex&, rowindex&    'Which day on the screen
Dim dx&, dy&                'Cursor position relative to the day

moDayIndex = -1
moDate = -1
moTimeSlot = -1
cx = X - MarginLeft
cy = Y - MarginTop
If (cx >= 0) And (cy >= 0) Then
    colindex = Int(cx / (DayWidth + DaySpacingX))
    rowindex = Int(cy / (DayHeight + DaySpacingY))
    If (colindex >= 0 And colindex <= 3) And (rowindex >= 0 And rowindex <= 1) Then
        If Not (rowindex = 0 And colindex = 3) Then       'Skip the gray area to the right of Wednesday
            dx = cx - (colindex * (DayWidth + DaySpacingX))
            dy = cy - (rowindex * (DayHeight + DaySpacingY))
            If (dx < DayWidth) And (dy < DayHeight) Then
                #If LRTB Then
                    moDayIndex = colindex + (rowindex * 3)
                #Else
                    moDayIndex = rowindex + (colindex * 2)
                #End If
                moDate = ViewStartDate + moDayIndex
                If (moDate >= ActiveDBInstance.ApptBitmap_StartDate) And (moDate <= (ActiveDBInstance.ApptBitmap_StartDate + ActiveDBInstance.ApptBitmap_Count - 1)) Then
                    moTimeSlot = Int((dy - DayFirstSlotOffsetY) / DayApptSlotHeight)
                    If (moTimeSlot > Appointment_NumSlotsUB) Or (moTimeSlot < 0) Then moTimeSlot = -1
                End If
            End If
        End If
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "MouseMoveCalc", Err
End Sub

'EHT=Standard
Sub SetShapeStyle(style As ScheduleShapeStyle)
On Error GoTo ERR_HANDLER

Select Case style
Case Style_Normal
    shpApptSelection.BorderColor = vbBlue
    lblApptSelection.Caption = ""
Case Style_New
    shpApptSelection.BorderColor = vbBlue
    lblApptSelection.Caption = "New"
Case Style_Move
    shpApptSelection.BorderColor = vbBlue
    lblApptSelection.Caption = "Move"
Case Style_Copy
    shpApptSelection.BorderColor = vbGreen
    lblApptSelection.Caption = "Copy"
Case Style_MoveAndCtrlCopy
    shpApptSelection.BorderColor = vbBlue
    lblApptSelection.Caption = "Move (hold Ctrl for Copy)"
Case Style_CopyForcedWithCtrl
    shpApptSelection.BorderColor = vbGreen
    lblApptSelection.Caption = "Copy (holding Ctrl)"
Case Style_ShowAppt
    shpApptSelection.BorderColor = vbRed
    lblApptSelection.Caption = ""
End Select
LastShapeStyle = style

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetShapeStyle", Err
End Sub

'EHT=Standard
Sub HideShape()
On Error GoTo ERR_HANDLER

shpApptSelection.Visible = False
lblApptSelection.Visible = False

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "HideShape", Err
End Sub

'EHT=Standard
Sub MoveShapeAndSetStyle(DayIndex&, TimeSlot&, NumSlots&, style As ScheduleShapeStyle)
On Error GoTo ERR_HANDLER

'Calculates the dimensions necessary to maintain the inner dimensions of the appointment rectangle
Dim w%, n1%, n2%
w = shpApptSelection.BorderWidth
n1 = (w - 1) \ 2
n2 = w - 1
shpApptSelection.Move ScheduleDayPositions(DayIndex).Left + DayApptsOffsetX - n1, _
         ScheduleDayPositions(DayIndex).Top + DayFirstSlotOffsetY + (TimeSlot * DayApptSlotHeight) - n1, _
         DayWidth - DayApptsOffsetX - DayMarginRight - 1 + n2, _
         (DayApptSlotHeight * NumSlots) + 1 + n2
shpApptSelection.Visible = True
lblApptSelection.Move ScheduleDayPositions(DayIndex).Left + DayApptsOffsetX + 1, _
         ScheduleDayPositions(DayIndex).Top + DayFirstSlotOffsetY + (TimeSlot * DayApptSlotHeight) + 1, _
         DayWidth - DayApptsOffsetX - DayMarginRight - 3, _
         (DayApptSlotHeight * NumSlots) - 1
lblApptSelection.Visible = True

SetShapeStyle style

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "MoveShapeAndSetStyle", Err
End Sub

'EHT=Standard
Sub MoveRedArrow(nt As Date)
On Error GoTo ERR_HANDLER

Dim tx&, ty&, tdi&, todaysdate&
tdi = CalcTodaysDayIndex(todaysdate)
If tdi <> TodaysDayIndex Then
    DrawSchedule      'Current day has changed, redraw (which will recalc the main TodaysDayIndex
Else
    If TodaysDayIndex = -1 Then
        imgCurTime.Visible = False
    Else
        tx = ScheduleDayPositions(TodaysDayIndex).Left
        ty = DayFirstSlotOffsetY + ((nt - Appointment_FirstSlotTime) / Appointment_SlotLength * DayApptSlotHeight)
        If (ty < 0) Or (ty > DayHeight) Then
            imgCurTime.Visible = False
        Else
            ty = ty + ScheduleDayPositions(TodaysDayIndex).Top - ((imgCurTime.Height - 1) / 2)
            tx = tx + DayApptsOffsetX - 13
            If imgCurTime.Left <> tx Then imgCurTime.Left = tx
            If imgCurTime.Top <> ty Then imgCurTime.Top = ty
            imgCurTime.Visible = True
        End If
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "MoveRedArrow", Err
End Sub

'EHT=Standard
Sub ShowDate(ByVal d As Date)
On Error GoTo ERR_HANDLER

Dim a&

If tmrFlashAppt.Enabled Then StopFlashAppt

'Make sure it is within bitmap range
a = Int(d)
If a < ActiveDBInstance.ApptBitmap_StartDate Then
    d = ActiveDBInstance.ApptBitmap_StartDate
ElseIf a >= (ActiveDBInstance.ApptBitmap_StartDate + ActiveDBInstance.ApptBitmap_Count) Then
    d = ActiveDBInstance.ApptBitmap_StartDate + ActiveDBInstance.ApptBitmap_Count - 1
End If

'Find week's Monday
For a = 1 To 7
    If Weekday(d) = vbMonday Then Exit For
    d = d - 1
Next a
a = Int(d)

ViewStartDate = a
shpApptSelection.Visible = False
lblApptSelection.Visible = False
If frmMain.CurTab = vSchedule Then DrawSchedule

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ShowDate", Err
End Sub

'EHT=Standard
Sub StopFlashAppt()
On Error GoTo ERR_HANDLER

tmrFlashAppt.Enabled = False
shpApptSelection.Visible = False

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "StopFlashAppt", Err
End Sub

