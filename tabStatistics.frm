VERSION 5.00
Begin VB.Form tabStatistics 
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
   Begin VB.CommandButton btnSaveLiveDataToSnapshot 
      Caption         =   "Create new Snapshot..."
      Height          =   360
      Left            =   840
      TabIndex        =   11
      Top             =   600
      Width           =   2295
   End
   Begin VB.ListBox lstSort 
      Height          =   300
      IntegralHeight  =   0   'False
      ItemData        =   "tabStatistics.frx":0000
      Left            =   7080
      List            =   "tabStatistics.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox pctDataView 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   1
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5895
   End
   Begin VB.PictureBox pctDataView 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   0
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   393
      TabIndex        =   8
      Top             =   360
      Width           =   5895
   End
   Begin VB.PictureBox pctControls 
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   1
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   4
      Top             =   1800
      Width           =   5655
      Begin VB.ComboBox cboSnapshots 
         Height          =   360
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   2535
      End
      Begin VB.CheckBox chkRememberSelection 
         Caption         =   "Remember?"
         Height          =   360
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblLiveIndicator 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4200
         TabIndex        =   7
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox pctControls 
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   0
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   377
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CheckBox chkRememberSelection 
         Caption         =   "Remember?"
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1455
      End
      Begin VB.ComboBox cboSnapshots 
         Height          =   360
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label lblLiveIndicator 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4200
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "tabStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabStatistics"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private SkipComboChangeEvent As Boolean

Private Const CurVersion As Byte = 3
Private Type typeSnapshot
    SnapshotVersion As Byte

    SnapshotDate As Date

    ReturnsCount_ByMonth(11) As Long
    ReturnsCountFeeGZ_ByMonth(11) As Long
    ReturnsCountMinGZ_ByMonth(11) As Long
    ReturnsTotalFee_ByMonth(11) As Long
    ReturnsTotalMin_ByMonth(11) As Long
    ReturnsUnpaidCount As Long
    ReturnsUnpaidTotalFee As Long
    XChgCount_ByMonth(11) As Long
    XChgTotalFee_ByMonth(11) As Long
    XChgUnpaidCount As Long
    XChgUnpaidTotalFee As Long
    BKTotalFee_ByMonth(11) As Long
    BKUnpaidTotalFee As Long

    NewCount As Long                'Marked complete, marked new, and meets NCThreshold
    NewTotal As Long
    NoShowCount As Long             'NOT marked complete, NOT marked extension, LYFlags marked complete
    NoShowTotal As Long
    NNTFCount As Long               'Marked NNTF
    NoChargeCount As Long           'Marked complete, fee = 0
    MailInCount As Long             'Marked complete, marked MI
    MailInTotal As Long
    DropOffCount As Long            'Marked complete, marked DO
    DropOffTotal As Long
    ExtensionsDoneCount As Long     'Marked complete, marked extension
    ExtensionsDoneTotal As Long
    ExtensionsCount As Long         'Marked extension
    StateListCount As Long          'Marked complete, statelist <> "" and statelist <> "CA"
    StateListTotal As Long
    EFiledCount As Long             'Marked complete, marked e-filed
    EFiledTotal As Long
    EmailAddressCount As Long       'Marked complete, email <> ""
    IPTECount As Long               'Marked complete, marked IPTE
    IPTETotal As Long

    '0-2 = Count, Total, Max
    CountTotalMaxData(4, 2) As Long
    'AGI_Data(2) As Long          'Skip IPTE
    'FedRef_Data(2) As Long       'Skip IPTE, FedResult >= 0
    'FedDue_Data(2) As Long       'Skip IPTE, FedResult < 0
    'StateRef_Data(2) As Long     'Skip IPTE, Skip NoState, StateResult >= 0
    'StateDue_Data(2) As Long     'Skip IPTE, Skip NoState, StateResult < 0

    '5 sectors, each with range start and end (inclusive), count, and total money
    BellCurveData(4, 3) As Long
End Type

Private Type typeTableDef
    pcthdc As Long
    OffsetX As Long
    OffsetY As Long
    RowHeight As Long
    ColumnWidth() As Long
    ColumnLeft() As Long
End Type

Private Const STAT_SnapshotFileExt = ".stat"
Private Const STAT_FontSize = 10
Private STAT_Font&
Private STAT_FontHeader&
Private BGBrush&

Private SnapshotSlot(1) As typeSnapshot
Private IsSnapshotLive(1) As Boolean

'EHT=Custom
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Standard
Private Function ITab_CreateGDIObjects() As Boolean
On Error GoTo ERR_HANDLER

STAT_FontHeader = CreateFont2(pctDataView(0).hdc, "Arial", STAT_FontSize, True, False, False, False)
STAT_Font = CreateFont2(pctDataView(0).hdc, "Arial", STAT_FontSize, False, False, False, False)
BGBrush = GetSysColorBrush(COLOR_WINDOW)

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CreateGDIObjects", Err
End Function

'EHT=Standard
Private Function ITab_InitializeAfterDBLoad() As Boolean
On Error GoTo ERR_HANDLER

chkRememberSelection(0).Enabled = ActiveDBInstance.IsWriteable
chkRememberSelection(1).Enabled = ActiveDBInstance.IsWriteable
btnSaveLiveDataToSnapshot.Enabled = ActiveDBInstance.IsWriteable

Dim a%, remember As Boolean
SkipComboChangeEvent = True
For a = 0 To 1
    remember = DB_GetSetting(ActiveDBInstance, "_Statistics-RememberSelection-" & a)
    chkRememberSelection(a).Value = (Not remember) + 1
    If remember Then
        PopulateComboWithSnapshotList a, DB_GetSetting(ActiveDBInstance, "_Statistics-LastView-" & a)
    Else
        PopulateComboWithSnapshotList a
        cboSnapshots(a).ListIndex = 0
    End If
Next a
SkipComboChangeEvent = False

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

PopulateComboWithSnapshotList 0
cboSnapshots_Click 0
PopulateComboWithSnapshotList 1
cboSnapshots_Click 1

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr pctDataView(0)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SetDefaultFocus", Err
End Sub

'EHT=Standard
Private Function ITab_SaveSettingsToDBBeforeClose() As Boolean
On Error GoTo ERR_HANDLER

Dim a%
For a = 0 To 1
    DB_SetSetting ActiveDBInstance, "_Statistics-RememberSelection-" & a, (chkRememberSelection(a).Value = 1), sBool
    DB_SetSetting ActiveDBInstance, "_Statistics-LastView-" & a, cboSnapshots(a).List(cboSnapshots(a).ListIndex), sStr
Next a

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SaveSettingsToDBBeforeClose", Err
End Function

'EHT=Standard
Private Function ITab_DestroyGDIObjects() As Boolean
On Error GoTo ERR_HANDLER

DeleteObject STAT_FontHeader
DeleteObject STAT_Font
DeleteObject BGBrush

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_DestroyGDIObjects", Err
End Function

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

Dim w&
w = (Me.ScaleWidth / 2) - 4
pctDataView(0).Move 0, pctControls(0).Height + 8, w, Me.ScaleHeight - pctControls(0).Height - 8
pctDataView(1).Move Me.ScaleWidth - w, pctControls(0).Height + 8, w, Me.ScaleHeight - pctControls(0).Height - 8
pctControls(0).Move pctDataView(0).Left + (pctDataView(0).Width / 2) - (pctControls(0).Width / 2), 0
pctControls(1).Move pctDataView(1).Left + (pctDataView(1).Width / 2) - (pctControls(1).Width / 2), 0
btnSaveLiveDataToSnapshot.Move (Me.ScaleWidth / 2) - (btnSaveLiveDataToSnapshot.Width / 2), 0

DisplayData 0
DisplayData 1
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
Private Sub btnSaveLiveDataToSnapshot_Click()
On Error GoTo ERR_HANDLER

'Save currently-displayed data to snapshot

Dim f$
f$ = InputBox("Enter title for new snapshot (without file extension):", , Format$(Now, "yyyy-mm") & " Custom Snapshot")
If f$ <> "" Then
    f$ = DataFilesPath & f$ & STAT_SnapshotFileExt
    If FileExists(f$) Then
        If MsgBox("File already exists. Replace?", vbYesNo Or vbDefaultButton2 Or vbExclamation) = vbNo Then
            Exit Sub
        End If
    End If
    If Not SaveLiveDataToSnapshotFile(f$) Then
        ShowErrorMsg "Failed to save live data to new snapshot file."
    End If
End If

'Refresh list
PopulateComboWithSnapshotList 0
cboSnapshots_Click 0
PopulateComboWithSnapshotList 1
cboSnapshots_Click 1

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSaveLiveDataToSnapshot_Click", Err
End Sub

'EHT=Standard
Sub cboSnapshots_Click(Index As Integer)
On Error GoTo ERR_HANDLER

Dim ls&

'Runs regardless of SkipComboChangeEvent >>>>>
    If cboSnapshots(Index).Tag = "" Then
        ls = -1
    Else
        ls = cboSnapshots(Index).Tag
    End If
    cboSnapshots(Index).Tag = cboSnapshots(Index).ListIndex
'<<<<<<<<<<<

If SkipComboChangeEvent Then Exit Sub
If cboSnapshots(Index).ListIndex < 0 Then Exit Sub

Select Case cboSnapshots(Index).ListIndex
Case 0
    'Calculate and display live data
    CalculateLiveDataAndPutIntoSlot Index
Case Else
    'Load snapshot file and display it
    LoadSnapshotFileIntoSlot cboSnapshots(Index).List(cboSnapshots(Index).ListIndex), Index
End Select

ITab_SetDefaultFocus

'Only set the changed flag if there would be a change in the settings saved to the DB
'This would only occur if a CBO changed that should be remembered (only allowed in write mode)
If ActiveDBInstance.IsWriteable Then
    If chkRememberSelection(Index).Value = vbChecked Then
        If cboSnapshots(Index).ListIndex <> ls Then
            frmMain.SetChangedFlagAndIndication
        End If
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "cboSnapshots_Click", Err
End Sub

'EHT=Standard
Private Sub chkRememberSelection_Click(Index As Integer)
On Error GoTo ERR_HANDLER

If Not ActiveDBInstance.IsWriteable Then Exit Sub
If SkipComboChangeEvent Then Exit Sub
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkRememberSelection_Click", Err
End Sub

'EHT=Cleanup2
Sub PopulateComboWithSnapshotList(Index As Integer, Optional ItemToSelect$)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

'If ItemToSelect$ is blank, selection won't change throughout the update

Dim t$, ps$, a&, os As Boolean

'Save the current selection
If ItemToSelect$ <> "" Then
    ps$ = ItemToSelect$
ElseIf cboSnapshots(Index).ListIndex >= 0 Then
    ps$ = cboSnapshots(Index).List(cboSnapshots(Index).ListIndex)
End If

cboSnapshots(Index).Clear
cboSnapshots(Index).AddItem "(Live Data)"
lstSort.Clear
t$ = Dir$(DataFilesPath & "*" & STAT_SnapshotFileExt)
Do Until t$ = ""
    'Add the filename portion only (skip the .stat)
    lstSort.AddItem Left$(t$, Len(t$) - Len(STAT_SnapshotFileExt))
    t$ = Dir$
Loop
For a = lstSort.ListCount - 1 To 0 Step -1
    cboSnapshots(Index).AddItem lstSort.List(a)
Next a

'Put the original selection back
For a = 0 To cboSnapshots(Index).ListCount - 1
    If cboSnapshots(Index).List(a) = ps$ Then
        os = SkipComboChangeEvent
        SkipComboChangeEvent = True
        cboSnapshots(Index).ListIndex = a
        SkipComboChangeEvent = os
        Exit For
    End If
Next a

CLEANUP: INCLEANUP = True
    If HASERROR Then cboSnapshots(Index).Clear

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PopulateComboWithSnapshotList", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Sub

'EHT=Standard
Private Function CalculateLiveData(ld As typeSnapshot) As Boolean
On Error GoTo ERR_HANDLER

Dim a&, m&, b As Boolean, ncft As Long, defstate$

defstate$ = DB_GetSetting(ActiveDBInstance, "GLOBAL_DefaultState")

'Clients
ncft = DB_GetSetting(ActiveDBInstance, "Prep fee threshold - new client SAF")
For a = 0 To 4
    ld.BellCurveData(a, 0) = DB_GetSetting(ActiveDBInstance, "Bell curve for statistics tab, range " & (a + 1) & " from")
    ld.BellCurveData(a, 1) = DB_GetSetting(ActiveDBInstance, "Bell curve for statistics tab, range " & (a + 1) & " to")
Next a
For a = 0 To ActiveDBInstance.Clients_Count - 1
    With ActiveDBInstance.Clients(a).c
        If Flag_IsSet(.Flags, CompletedReturn) Then
            If .CompletionDate = NullLong Then Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", CompletionDate cannot be null"
            m = Month(.CompletionDate) - 1
            Inc ld.ReturnsCount_ByMonth(m)
            If .PrepFee = NullLong Then Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", PrepFee cannot be null"
            IncBy ld.ReturnsTotalFee_ByMonth(m), .PrepFee
            If .MinutesToComplete = NullLong Then Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", MinutesToComplete cannot be null"
            IncBy ld.ReturnsTotalMin_ByMonth(m), .MinutesToComplete
            If .PrepFee > 0 Then Inc ld.ReturnsCountFeeGZ_ByMonth(m)
            If .MinutesToComplete > 0 Then Inc ld.ReturnsCountMinGZ_ByMonth(m)
            If .MoneyOwed <> NullLong Then
                Inc ld.ReturnsUnpaidCount
                IncBy ld.ReturnsUnpaidTotalFee, .MoneyOwed
            End If

            If Flag_IsSet(.Flags, NewClient) Then
                If .PrepFee >= ncft Then
                    Inc ld.NewCount
                    IncBy ld.NewTotal, .PrepFee
                End If
            End If
            If .PrepFee = 0 Then
                Inc ld.NoChargeCount
            End If
            If Flag_IsSet(.Flags, MailedIn) Then
                Inc ld.MailInCount
                IncBy ld.MailInTotal, .PrepFee
            End If
            If Flag_IsSet(.Flags, DroppedOff) Then
                Inc ld.DropOffCount
                IncBy ld.DropOffTotal, .PrepFee
            End If
            If Len(.StateList) > 0 Then
                If .StateList <> defstate$ Then
                    Inc ld.StateListCount
                    IncBy ld.StateListTotal, .PrepFee
                End If
            End If
            If Flag_IsSet(.Flags, EFiled) Then
                Inc ld.EFiledCount
                IncBy ld.EFiledTotal, .PrepFee
            End If
            If .Person1.Email <> "" Or .Person2.Email <> "" Then
                Inc ld.EmailAddressCount
            End If
            If Flag_IsSet(.Flags, IncPtnrTrustEstate) Then
                Inc ld.IPTECount
                IncBy ld.IPTETotal, .PrepFee
            Else
                Inc ld.CountTotalMaxData(0, 0)
                If .ResultAGI = NullLong Then Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", ResultAGI cannot be null"
                IncBy ld.CountTotalMaxData(0, 1), .ResultAGI
                If .ResultAGI > ld.CountTotalMaxData(0, 2) Then ld.CountTotalMaxData(0, 2) = .ResultAGI

                If .ResultFederal = NullLong Then
                    Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", ResultFederal cannot be null"
                ElseIf .ResultFederal < 0 Then
                    Inc ld.CountTotalMaxData(2, 0)
                    IncBy ld.CountTotalMaxData(2, 1), (-.ResultFederal)
                    If (-.ResultFederal) > ld.CountTotalMaxData(2, 2) Then ld.CountTotalMaxData(2, 2) = (-.ResultFederal)
                Else
                    Inc ld.CountTotalMaxData(1, 0)
                    IncBy ld.CountTotalMaxData(1, 1), .ResultFederal
                    If .ResultFederal > ld.CountTotalMaxData(1, 2) Then ld.CountTotalMaxData(1, 2) = .ResultFederal
                End If

                If Len(.StateList) > 0 Then
                    If .ResultState = NullLong Then
                        Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", ResultState cannot be null"
                    ElseIf .ResultState < 0 Then
                        Inc ld.CountTotalMaxData(4, 0)
                        IncBy ld.CountTotalMaxData(4, 1), (-.ResultState)
                        If (-.ResultState) > ld.CountTotalMaxData(4, 2) Then ld.CountTotalMaxData(4, 2) = (-.ResultState)
                    Else
                        Inc ld.CountTotalMaxData(3, 0)
                        IncBy ld.CountTotalMaxData(3, 1), .ResultState
                        If .ResultState > ld.CountTotalMaxData(3, 2) Then ld.CountTotalMaxData(3, 2) = .ResultState
                    End If
                End If
            End If

            For m = 0 To 4
                b = False
                If ld.BellCurveData(m, 0) <> NullLong Then
                    b = (.PrepFee >= ld.BellCurveData(m, 0))
                End If
                If ld.BellCurveData(m, 1) <> NullLong Then
                    b = b And (.PrepFee <= ld.BellCurveData(m, 1))
                End If
                If b Then
                    Inc ld.BellCurveData(m, 2)
                    IncBy ld.BellCurveData(m, 3), .PrepFee
                End If
            Next m
        End If
        If (Not Flag_IsSet(.Flags, CompletedReturn)) And (Not Flag_IsSet(.Flags, Extension)) And Flag_IsSet(.LastYear_Flags, CompletedReturn) Then
            Inc ld.NoShowCount
            If .LastYear_PrepFee = NullLong Then Err.Raise 1, , FormatClientName(fLog, ActiveDBInstance.Clients(a).c) & ", LYPrepFee cannot be null"
            IncBy ld.NoShowTotal, .LastYear_PrepFee
        End If
        If Flag_IsSet(.Flags, NoNeedToFile) Then
            Inc ld.NNTFCount
        End If
        If Flag_IsSet(.Flags, Extension) Then
            Inc ld.ExtensionsCount
            If Flag_IsSet(.Flags, CompletedReturn) Then
                Inc ld.ExtensionsDoneCount
                IncBy ld.ExtensionsDoneTotal, .PrepFee
            End If
        End If
    End With
Next a

'Extra Charges
For a = 0 To ActiveDBInstance.ExtraCharges_Count - 1
    With ActiveDBInstance.ExtraCharges(a)
        m = Month(.CompletionDate) - 1
        Inc ld.XChgCount_ByMonth(m)
        If .PrepFee = NullLong Then Err.Raise 1, , "ExChgID#" & a & " (" & .ClientName & "), PrepFee cannot be null"
        IncBy ld.XChgTotalFee_ByMonth(m), .PrepFee
        If .MoneyOwed <> NullLong Then
            Inc ld.XChgUnpaidCount
            IncBy ld.XChgUnpaidTotalFee, .MoneyOwed
        End If
    End With
Next a

'Bookkeeping
For a = 0 To ActiveDBInstance.Bookkeeping_Count - 1
    For m = 0 To 11
        With ActiveDBInstance.Bookkeeping(a).Months(m)
            If .PrepFee <> NullLong Then IncBy ld.BKTotalFee_ByMonth(m), .PrepFee
            If .MoneyOwed <> NullLong Then IncBy ld.BKUnpaidTotalFee, .MoneyOwed
        End With
    Next m
Next a

'Finalize
ld.SnapshotVersion = CurVersion
ld.SnapshotDate = Now
CalculateLiveData = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CalculateLiveData", Err
End Function

'EHT=Cleanup1
Sub CalculateLiveDataAndPutIntoSlot(SlotIndex As Integer)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim ld As typeSnapshot
SnapshotSlot(SlotIndex) = ld            'Clear it out first; so if there is any error, the slot is left empty
IsSnapshotLive(SlotIndex) = False

If CalculateLiveData(ld) Then
    'Succeeded, so put it into the slot and display it
    SnapshotSlot(SlotIndex) = ld
    IsSnapshotLive(SlotIndex) = True
End If

CLEANUP: INCLEANUP = True
    DisplayData SlotIndex       'Display what we have, regardless of succeess

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CalculateLiveDataAndPutIntoSlot", Err, INCLEANUP: Resume CLEANUP
End Sub

'EHT=Cleanup1
Sub DisplayData(SlotIndex As Integer)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim a&, b&, b2&, secX&, secY&, y2&, t$
Dim pcthdc&, tbl As typeTableDef
Dim ct&(4, 7), NumReturnsCompleted&, TotalReturnIncome&

pctDataView(SlotIndex).Cls

pcthdc = pctDataView(SlotIndex).hdc

With SnapshotSlot(SlotIndex)
    If .SnapshotDate = 0 Then Exit Sub

    tbl.pcthdc = pcthdc
    tbl.OffsetX = 0
    tbl.OffsetY = 0
    tbl.RowHeight = 17
    ReDim tbl.ColumnWidth(15)
    tbl.ColumnWidth(0) = 80
    tbl.ColumnWidth(1) = 30
    tbl.ColumnWidth(2) = 60
    tbl.ColumnWidth(3) = 40
    tbl.ColumnWidth(4) = 60
    tbl.ColumnWidth(5) = 40
    tbl.ColumnWidth(6) = 30
    tbl.ColumnWidth(7) = 30
    tbl.ColumnWidth(8) = 30
    tbl.ColumnWidth(9) = 30
    tbl.ColumnWidth(10) = 30
    tbl.ColumnWidth(11) = 75
    tbl.ColumnWidth(12) = 30
    tbl.ColumnWidth(13) = 80
    tbl.ColumnWidth(14) = 60
    tbl.ColumnWidth(15) = 60
    InitTable tbl

    '***************************************************************************************
    secX = 0: secY = 0
    FillArea tbl, secX + 0, secY + 0, 12, 21
    'Horizontal lines
    DrawBorders tbl, secX + 0, secY + 0, 12, 21, True, True, True, True
    DrawBorders tbl, secX + 0, secY + 1, 12, 1, False, True, False, False
    DrawBorders tbl, secX + 0, secY + 6, 12, 1, True, True, False, False
    DrawBorders tbl, secX + 0, secY + 7, 12, 1, True, True, False, False
    DrawBorders tbl, secX + 0, secY + 8, 12, 1, True, True, False, False
    DrawBorders tbl, secX + 0, secY + 17, 12, 1, True, False, False, False
    DrawBorders tbl, secX + 0, secY + 18, 12, 1, True, False, False, False
    DrawBorders tbl, secX + 0, secY + 19, 12, 1, True, False, False, False
    DrawBorders tbl, secX + 0, secY + 20, 12, 1, True, False, False, False
    'Vertical lines
    DrawBorders tbl, secX + 0, secY + 0, 1, 21, False, False, False, True
    DrawBorders tbl, secX + 5, secY + 0, 1, 21, False, False, False, True
    DrawBorders tbl, secX + 8, secY + 0, 1, 21, False, False, False, True
    DrawBorders tbl, secX + 10, secY + 0, 1, 21, False, False, False, True

    SetTextColor pcthdc, vbBlack
    SelectObject pcthdc, STAT_FontHeader
    If .SnapshotDate <> 0 Then DrawContent tbl, secX + 0, secY + 0, 1, 2, Format$(.SnapshotDate, "m/dd/yyyy")
    DrawContent tbl, secX + 1, secY + 0, 5, 1, "Returns"
    DrawContent tbl, secX + 6, secY + 0, 3, 1, "XChg"
    DrawContent tbl, secX + 9, secY + 0, 2, 1, "BK"
    DrawContent tbl, secX + 11, secY + 0, 1, 1, "Total"
    DrawContent tbl, secX + 11, secY + 1, 1, 1, "Income"

    SelectObject pcthdc, STAT_Font
    DrawContent tbl, secX + 1, secY + 1, 1, 1, "#"
    DrawContent tbl, secX + 2, secY + 1, 1, 1, "Total"
    DrawContent tbl, secX + 3, secY + 1, 1, 1, "Avg"
    DrawContent tbl, secX + 4, secY + 1, 1, 1, "Minutes"
    DrawContent tbl, secX + 5, secY + 1, 1, 1, "Avg"
    DrawContent tbl, secX + 6, secY + 1, 1, 1, "#"
    DrawContent tbl, secX + 7, secY + 1, 2, 1, "Total"
    DrawContent tbl, secX + 9, secY + 1, 2, 1, "Total"

    'ct(0,_)    Jan,Feb
    'ct(1,_)    Jan,Feb,Mar
    'ct(2,_)    Jan,Feb,Mar,Apr
    'ct(3,_)    May-Dec
    'ct(4,_)    Entire year
    For a = 0 To 11
        If a < 4 Then
            y2 = a + 2      'Offset by the 2 rows of the header

            b2 = a - 1
            If b2 < 0 Then b2 = 0
            For b = b2 To 2
                'Increment Jan-Feb, Jan-Mar, and Jan-Apr subtotals
                IncBy ct(b, 0), .ReturnsCount_ByMonth(a)
                IncBy ct(b, 1), .ReturnsCountFeeGZ_ByMonth(a)
                IncBy ct(b, 2), .ReturnsCountMinGZ_ByMonth(a)
                IncBy ct(b, 3), .ReturnsTotalFee_ByMonth(a)
                IncBy ct(b, 4), .ReturnsTotalMin_ByMonth(a)
                IncBy ct(b, 5), .XChgCount_ByMonth(a)
                IncBy ct(b, 6), .XChgTotalFee_ByMonth(a)
                IncBy ct(b, 7), .BKTotalFee_ByMonth(a)
            Next b

        Else
            y2 = a + 5      'Offset by the 2 header rows, plus the 3 subtotal rows

            'Increment May-Dec subtotals
            IncBy ct(3, 0), .ReturnsCount_ByMonth(a)
            IncBy ct(3, 1), .ReturnsCountFeeGZ_ByMonth(a)
            IncBy ct(3, 2), .ReturnsCountMinGZ_ByMonth(a)
            IncBy ct(3, 3), .ReturnsTotalFee_ByMonth(a)
            IncBy ct(3, 4), .ReturnsTotalMin_ByMonth(a)
            IncBy ct(3, 5), .XChgCount_ByMonth(a)
            IncBy ct(3, 6), .XChgTotalFee_ByMonth(a)
            IncBy ct(3, 7), .BKTotalFee_ByMonth(a)
        End If

        'Increment year totals
        IncBy ct(4, 0), .ReturnsCount_ByMonth(a)
        IncBy ct(4, 1), .ReturnsCountFeeGZ_ByMonth(a)
        IncBy ct(4, 2), .ReturnsCountMinGZ_ByMonth(a)
        IncBy ct(4, 3), .ReturnsTotalFee_ByMonth(a)
        IncBy ct(4, 4), .ReturnsTotalMin_ByMonth(a)
        IncBy ct(4, 5), .XChgCount_ByMonth(a)
        IncBy ct(4, 6), .XChgTotalFee_ByMonth(a)
        IncBy ct(4, 7), .BKTotalFee_ByMonth(a)

        'Draw that month's line
        DrawContent tbl, secX + 0, secY + y2, 1, 1, "   " & MonthName(a + 1), DT_LEFT
        DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(.ReturnsCount_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(.ReturnsTotalFee_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtAvgTotalOrBlank(.ReturnsTotalFee_ByMonth(a), .ReturnsCountFeeGZ_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtCountOrBlank(.ReturnsTotalMin_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 5, secY + y2, 1, 1, FmtAvgCountOrBlank(.ReturnsTotalMin_ByMonth(a), .ReturnsCountMinGZ_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(.XChgCount_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(.XChgTotalFee_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(.BKTotalFee_ByMonth(a)), DT_RIGHT
        DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(.ReturnsTotalFee_ByMonth(a) + .XChgTotalFee_ByMonth(a) + .BKTotalFee_ByMonth(a)), DT_RIGHT
    Next a
    NumReturnsCompleted = ct(4, 0)
    TotalReturnIncome = ct(4, 3)

    SelectObject pcthdc, STAT_FontHeader
    SetTextColor pcthdc, &HFF6633

    y2 = 6
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "Jan-Feb", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(ct(0, 0)), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(ct(0, 3)), DT_RIGHT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtAvgTotalOrBlank(ct(0, 3), ct(0, 1)), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtCountOrBlank(ct(0, 4)), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 1, 1, FmtAvgCountOrBlank(ct(0, 4), ct(0, 2)), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(ct(0, 5)), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(ct(0, 6)), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(ct(0, 7)), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(ct(0, 3) + ct(0, 6) + ct(0, 7)), DT_RIGHT

    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "Jan-Mar", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(ct(1, 0)), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(ct(1, 3)), DT_RIGHT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtAvgTotalOrBlank(ct(1, 3), ct(1, 1)), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtCountOrBlank(ct(1, 4)), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 1, 1, FmtAvgCountOrBlank(ct(1, 4), ct(1, 2)), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(ct(1, 5)), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(ct(1, 6)), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(ct(1, 7)), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(ct(1, 3) + ct(1, 6) + ct(1, 7)), DT_RIGHT

    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "Jan-Apr", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(ct(2, 0)), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(ct(2, 3)), DT_RIGHT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtAvgTotalOrBlank(ct(2, 3), ct(2, 1)), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtCountOrBlank(ct(2, 4)), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 1, 1, FmtAvgCountOrBlank(ct(2, 4), ct(2, 2)), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(ct(2, 5)), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(ct(2, 6)), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(ct(2, 7)), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(ct(2, 3) + ct(2, 6) + ct(2, 7)), DT_RIGHT

    y2 = 17
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "May-Dec", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(ct(3, 0)), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(ct(3, 3)), DT_RIGHT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtAvgTotalOrBlank(ct(3, 3), ct(3, 1)), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtCountOrBlank(ct(3, 4)), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 1, 1, FmtAvgCountOrBlank(ct(3, 4), ct(3, 2)), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(ct(3, 5)), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(ct(3, 6)), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(ct(3, 7)), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(ct(3, 3) + ct(3, 6) + ct(3, 7)), DT_RIGHT

    SetTextColor pcthdc, vbBlack
    y2 = 18
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "Year", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(ct(4, 0)), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(ct(4, 3)), DT_RIGHT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtAvgTotalOrBlank(ct(4, 3), ct(4, 1)), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtCountOrBlank(ct(4, 4)), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 1, 1, FmtAvgCountOrBlank(ct(4, 4), ct(4, 2)), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(ct(4, 5)), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(ct(4, 6)), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(ct(4, 7)), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(ct(4, 3) + ct(4, 6) + ct(4, 7)), DT_RIGHT

    SetTextColor pcthdc, vbRed
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "Owed", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(.ReturnsUnpaidCount), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(-.ReturnsUnpaidTotalFee), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(.XChgUnpaidCount), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(-.XChgUnpaidTotalFee), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(-.BKUnpaidTotalFee), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(-.ReturnsUnpaidTotalFee - .XChgUnpaidTotalFee - .BKUnpaidTotalFee), DT_RIGHT

    SetTextColor pcthdc, &HC000&
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 1, 1, "Received", DT_LEFT
    DrawContent tbl, secX + 1, secY + y2, 1, 1, FmtCountOrBlank(ct(4, 0) - .ReturnsUnpaidCount), DT_RIGHT
    DrawContent tbl, secX + 2, secY + y2, 1, 1, FmtTotalOrBlank(ct(4, 3) - .ReturnsUnpaidTotalFee), DT_RIGHT
    DrawContent tbl, secX + 6, secY + y2, 1, 1, FmtCountOrBlank(ct(4, 5) - .XChgUnpaidCount), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtTotalOrBlank(ct(4, 6) - .XChgUnpaidTotalFee), DT_RIGHT
    DrawContent tbl, secX + 9, secY + y2, 2, 1, FmtTotalOrBlank(ct(4, 7) - .BKUnpaidTotalFee), DT_RIGHT
    DrawContent tbl, secX + 11, secY + y2, 1, 1, FmtTotalOrBlank(ct(4, 3) + ct(4, 6) + ct(4, 7) - .ReturnsUnpaidTotalFee - .XChgUnpaidTotalFee - .BKUnpaidTotalFee), DT_RIGHT



    '***************************************************************************************
    secX = 0: secY = 22
    FillArea tbl, secX + 3, secY + 0, 6, 1
    FillArea tbl, secX + 0, secY + 1, 9, 12
    DrawBorders tbl, secX + 3, secY + 0, 2, 1, True, False, True, False
    DrawBorders tbl, secX + 5, secY + 0, 4, 1, True, False, True, True
    DrawBorders tbl, secX + 0, secY + 1, 3, 12, True, True, True, False
    DrawBorders tbl, secX + 3, secY + 1, 2, 12, True, True, True, False
    DrawBorders tbl, secX + 5, secY + 1, 4, 12, True, True, True, True
    SetTextColor pcthdc, vbBlack

    SelectObject pcthdc, STAT_FontHeader
    DrawContent tbl, secX + 5, secY + 0, 2, 1, "Total"

    SelectObject pcthdc, STAT_Font
    DrawContent tbl, secX + 3, secY + 0, 1, 1, "#"
    DrawContent tbl, secX + 4, secY + 0, 1, 1, "%"
    DrawContent tbl, secX + 7, secY + 0, 2, 1, "%"

    y2 = 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "New - SAF", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.NewCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.NewCount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.NewTotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.NewTotal, TotalReturnIncome), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "No-Shows", DT_LEFT
    If .NoShowCount > 0 Then DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.NoShowCount), DT_RIGHT
    If .NoShowTotal > 0 Then DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.NoShowTotal), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Mail-In", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.MailInCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.MailInCount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.MailInTotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.MailInTotal, TotalReturnIncome), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Drop-Off", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.DropOffCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.DropOffCount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.DropOffTotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.DropOffTotal, TotalReturnIncome), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "NNTF", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.NNTFCount), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Returns @ No Charge", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.NoChargeCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.NoChargeCount, NumReturnsCompleted), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Extensions - Total", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.ExtensionsCount), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Extensions - Done", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.ExtensionsDoneCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.ExtensionsDoneCount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.ExtensionsDoneTotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.ExtensionsDoneTotal, TotalReturnIncome), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Other State Returns", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.StateListCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.StateListCount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.StateListTotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.StateListTotal, TotalReturnIncome), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "E-Filed Returns", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.EFiledCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.EFiledCount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.EFiledTotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.EFiledTotal, TotalReturnIncome), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "Email Addresses", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.EmailAddressCount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.EmailAddressCount, NumReturnsCompleted), DT_RIGHT
    y2 = y2 + 1
    DrawContent tbl, secX + 0, secY + y2, 3, 1, "IPTE Returns", DT_LEFT
    DrawContent tbl, secX + 3, secY + y2, 1, 1, FmtCount(.IPTECount), DT_RIGHT
    DrawContent tbl, secX + 4, secY + y2, 1, 1, FmtPercent(.IPTECount, NumReturnsCompleted), DT_RIGHT
    DrawContent tbl, secX + 5, secY + y2, 2, 1, FmtTotal(.IPTETotal), DT_RIGHT
    DrawContent tbl, secX + 7, secY + y2, 2, 1, FmtPercent(.IPTETotal, TotalReturnIncome), DT_RIGHT







    '***************************************************************************************
    secX = 0: secY = 36
    FillArea tbl, secX + 1, secY, 6, 1
    FillArea tbl, secX + 0, secY + 1, 7, 5
    DrawBorders tbl, secX + 1, secY + 0, 6, 1, True, False, True, True
    DrawBorders tbl, secX + 0, secY + 1, 1, 5, True, True, True, False
    DrawBorders tbl, secX + 1, secY + 1, 6, 5, True, True, True, True
    SetTextColor pcthdc, vbBlack

    SelectObject pcthdc, STAT_FontHeader
    DrawContent tbl, secX + 2, secY + 0, 2, 1, "Total"
    DrawContent tbl, secX + 4, secY + 0, 1, 1, "Average"
    DrawContent tbl, secX + 5, secY + 0, 2, 1, "Max"

    SelectObject pcthdc, STAT_Font
    DrawContent tbl, secX + 1, secY + 0, 1, 1, "#"

    For a = 0 To 4
        t$ = Choose(a + 1, "AGI", "FedRef", "FedDue", "StateRef", "StateDue")
        DrawContent tbl, secX + 0, secY + 1 + a, 1, 1, t$, DT_LEFT
        DrawContent tbl, secX + 1, secY + 1 + a, 1, 1, FmtCountOrBlank(.CountTotalMaxData(a, 0)), DT_RIGHT
        DrawContent tbl, secX + 2, secY + 1 + a, 2, 1, FmtTotalOrBlank(.CountTotalMaxData(a, 1)), DT_RIGHT
        DrawContent tbl, secX + 4, secY + 1 + a, 1, 1, FmtAvgTotalOrBlank(.CountTotalMaxData(a, 1), .CountTotalMaxData(a, 0)), DT_RIGHT
        DrawContent tbl, secX + 5, secY + 1 + a, 2, 1, FmtTotalOrBlank(.CountTotalMaxData(a, 2)), DT_RIGHT
    Next a






    '***************************************************************************************
    secX = 8: secY = 36
    FillArea tbl, secX + 2, secY + 0, 2, 1
    FillArea tbl, secX + 0, secY + 1, 4, 5
    DrawBorders tbl, secX + 2, secY + 0, 2, 1, True, False, True, True
    DrawBorders tbl, secX + 0, secY + 1, 2, 5, True, True, True, False
    DrawBorders tbl, secX + 2, secY + 1, 2, 5, True, True, True, True
    SetTextColor pcthdc, vbBlack

    SelectObject pcthdc, STAT_FontHeader
    DrawContent tbl, secX + 3, secY + 0, 1, 1, "Total"

    SelectObject pcthdc, STAT_Font
    DrawContent tbl, secX + 2, secY + 0, 1, 1, "#"

    For a = 0 To 4
        If .BellCurveData(a, 0) = NullLong Then
            t$ = "<= " & FmtCount(.BellCurveData(a, 1))
        ElseIf .BellCurveData(a, 1) = NullLong Then
            t$ = "> " & FmtCount(.BellCurveData(a, 0) - 1)
        Else
            t$ = FmtCount(.BellCurveData(a, 0)) & " - " & FmtCount(.BellCurveData(a, 1))
        End If
        DrawContent tbl, secX + 0, secY + 1 + a, 2, 1, t$, DT_RIGHT
        DrawContent tbl, secX + 2, secY + 1 + a, 1, 1, FmtCountOrBlank(.BellCurveData(a, 2)), DT_RIGHT
        DrawContent tbl, secX + 3, secY + 1 + a, 1, 1, FmtTotalOrBlank(.BellCurveData(a, 3)), DT_RIGHT
    Next a
End With

If IsSnapshotLive(SlotIndex) Then
    If SnapshotSlot(SlotIndex).SnapshotDate = 0 Then
        lblLiveIndicator(SlotIndex).Caption = ""
    Else
        lblLiveIndicator(SlotIndex).Caption = "Live Data"
    End If
    lblLiveIndicator(SlotIndex).ForeColor = &HC000&
Else
    lblLiveIndicator(SlotIndex).Caption = "Snapshot"
    lblLiveIndicator(SlotIndex).ForeColor = vbRed
End If

CLEANUP: INCLEANUP = True
    pctDataView(SlotIndex).Refresh

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DisplayData", Err, INCLEANUP: Resume CLEANUP
End Sub

'EHT=Cleanup1
Sub LoadSnapshotFileIntoSlot(ByVal filetoload$, SlotIndex As Integer)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim nd As typeSnapshot
SnapshotSlot(SlotIndex) = nd            'Clear it out first; that way if there is an error, the slot is left empty
IsSnapshotLive(SlotIndex) = False

filetoload$ = DataFilesPath & filetoload$ & STAT_SnapshotFileExt
If Not FileExists(filetoload$) Then
    ShowErrorMsg "File does not exist."
Else
    Dim fh As CMNMOD_CFileHandler
    Set fh = OpenFile(filetoload$, mBinary_Input)
    Get #fh.FileNum, , nd
    fh.CloseFile: Set fh = Nothing
    If nd.SnapshotVersion <> CurVersion Then
        If MsgBox("The selected snapshot has been saved with an older version of the program. Some data may show incorrectly. Load anyway?", vbQuestion Or vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    SnapshotSlot(SlotIndex) = nd            'Put the new data into the slot
    IsSnapshotLive(SlotIndex) = False
End If

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing
    DisplayData SlotIndex       'Display what we have, regardless of success

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "LoadSnapshotFileIntoSlot", Err, INCLEANUP: Resume CLEANUP
End Sub

'EHT=Cleanup1
Function SaveLiveDataToSnapshotFile(filetosaveto$) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim ld As typeSnapshot
If CalculateLiveData(ld) Then
    Dim fh As CMNMOD_CFileHandler
    Set fh = OpenFile(filetosaveto$, mBinary_Output)
        Put #fh.FileNum, , ld
    fh.CloseFile: Set fh = Nothing
End If

''Reload the active data view and select the new entry
'PopulateComboWithSnapshotList SlotIndex, filetosaveto$
''Reload the other one without changing the selection
'If SlotIndex = 0 Then
'    PopulateComboWithSnapshotList 1
'Else
'    PopulateComboWithSnapshotList 0
'End If

SaveLiveDataToSnapshotFile = True

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SaveLiveDataToSnapshotFile", Err, INCLEANUP: Resume CLEANUP
End Function

'EHT=Standard
Sub CreateAutoSnapshotIfNewMonth()
On Error GoTo ERR_HANDLER

'If this is a new month, and the user forgot to create a snapshot, create one automatically

Dim MostRecentSnapshot$, CurDBCalendarYear%, MostRecentSnapshotMonth%
Dim CurMonth%, CurYear%
Dim lc%, a%, b$, f$, m$

'Get current database's calendar year (calendar year is the year in the filename *plus one*)
CurDBCalendarYear = Val(Mid(ActiveDBInstance.FullPath_DB, Len(ActiveDBInstance.FullPath_DB) - 7, 4)) + 1
If CurDBCalendarYear = 0 Then Err.Raise 1, , "Unable to determine current database year."

'Find the most recent snapshot whose name begins with the 4-digit code of the current database
lc = cboSnapshots(0).ListCount
For a = 0 To lc - 1
    MostRecentSnapshot$ = cboSnapshots(0).List(a)
    If MostRecentSnapshot$ Like (Format(CurDBCalendarYear, "0000") & "-?? *(auto)") Then
        MostRecentSnapshotMonth = Val(Mid$(MostRecentSnapshot$, 6, 2))
        If MostRecentSnapshotMonth >= 1 And MostRecentSnapshotMonth <= 12 Then Exit For
    End If
Next a

'Determine any missing months between that snapshot and the present month
'(assuming present month is in the same year as the database)
CurYear = Year(Date)
If CurYear < CurDBCalendarYear Then Exit Sub        'Opening a future database? This is weird; get out now!
If CurYear > CurDBCalendarYear Then
    CurMonth = 13                           'Opening a past database? Use 13 to trigger the rest of the code to create snapshots X through 12
Else
    CurMonth = Month(Date)
End If

'Create any missing snapshots, using live data (we know live data is good for this purpose,
'because the fact that the snapshot is missing in the first place tells us that the program
'has not been opened since that snapshot was made... thus, no data could have changed either.)
For a = MostRecentSnapshotMonth + 1 To CurMonth - 1
    b$ = Format(DateSerial(CurDBCalendarYear, a, 1), "yyyy-mm mmmm") & " (auto)"
    f$ = DataFilesPath & b$ & STAT_SnapshotFileExt
    If Not FileExists(f$) Then
        If SaveLiveDataToSnapshotFile(f$) Then
            m$ = m$ & vbCrLf & b$
        End If
    End If
Next a
If Len(m$) > 0 Then
    ShowInfoMsg "Automatically created the following snapshots using live data as of this moment:" & m$
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CreateAutoSnapshotIfNewMonth", Err
End Sub

'EHT=Custom
Private Sub InitTable(t As typeTableDef)
Dim a&, ub&
ub = UBound(t.ColumnWidth)
ReDim t.ColumnLeft(ub)
For a = 1 To ub
    t.ColumnLeft(a) = t.ColumnLeft(a - 1) + t.ColumnWidth(a - 1) - 1
Next a
End Sub

'EHT=Custom
Private Sub DrawBorders(table As typeTableDef, c&, r&, nc&, nr&, bt As Boolean, bb As Boolean, bl As Boolean, br As Boolean)
Dim dr As RECT
dr.Top = table.OffsetY + r * (table.RowHeight - 1)
dr.Left = table.OffsetX + table.ColumnLeft(c)
dr.Bottom = table.OffsetY + (r + nr) * (table.RowHeight - 1)
dr.Right = table.OffsetX + table.ColumnLeft(c + nc - 1) + table.ColumnWidth(c + nc - 1) - 1
If bt Then
    MoveToEx table.pcthdc, dr.Left, dr.Top, 0
    LineTo table.pcthdc, dr.Right + 1, dr.Top
End If
If bb Then
    MoveToEx table.pcthdc, dr.Left, dr.Bottom, 0
    LineTo table.pcthdc, dr.Right + 1, dr.Bottom
End If
If bl Then
    MoveToEx table.pcthdc, dr.Left, dr.Top, 0
    LineTo table.pcthdc, dr.Left, dr.Bottom + 1
End If
If br Then
    MoveToEx table.pcthdc, dr.Right, dr.Top, 0
    LineTo table.pcthdc, dr.Right, dr.Bottom + 1
End If
End Sub

'EHT=Custom
Private Sub DrawContent(table As typeTableDef, c&, r&, nc&, nr&, s$, Optional halign& = DT_CENTER, Optional valign& = DT_VCENTER)
Dim dr As RECT
dr.Top = 1 + table.OffsetY + r * (table.RowHeight - 1)
dr.Left = 2 + table.OffsetX + table.ColumnLeft(c)
dr.Bottom = table.OffsetY + (r + nr) * (table.RowHeight - 1)
dr.Right = -1 + table.OffsetX + table.ColumnLeft(c + nc - 1) + table.ColumnWidth(c + nc - 1) - 1
DrawText table.pcthdc, s$, Len(s$), dr, halign Or valign Or DT_SINGLELINE
End Sub

'EHT=Custom
Private Sub FillArea(table As typeTableDef, c&, r&, nc&, nr&)
Dim dr As RECT
dr.Top = table.OffsetY + r * (table.RowHeight - 1)
dr.Left = table.OffsetX + table.ColumnLeft(c)
dr.Bottom = table.OffsetY + (r + nr) * (table.RowHeight - 1)
dr.Right = table.OffsetX + table.ColumnLeft(c + nc - 1) + table.ColumnWidth(c + nc - 1) - 1
FillRect table.pcthdc, dr, BGBrush
End Sub

'EHT=Custom
Private Function FmtAvgCountOrBlank$(l1&, l2&)
If l2 <> 0 Then
    FmtAvgCountOrBlank = Format$(l1 / l2, "#,##0")
End If
End Function

'EHT=Custom
Private Function FmtAvgTotalOrBlank$(l1&, l2&)
If l2 <> 0 Then
    FmtAvgTotalOrBlank = Format$(l1 / l2, "$#,##0")
End If
End Function

'EHT=Custom
Private Function FmtCount$(l&)
FmtCount = Format$(l, "#,##0")
End Function

'EHT=Custom
Private Function FmtCountOrBlank$(l&)
If l <> 0 Then FmtCountOrBlank = Format$(l, "#,##0")
End Function

'EHT=Custom
Private Function FmtPercent$(l1&, l2&)
Dim p#
If l2 = 0 Then
    p = 0
Else
    p = l1 / l2
End If
FmtPercent = Format$(p, "0%")
End Function

'EHT=Custom
Private Function FmtPercentOrBlank$(l1&, l2&)
Dim p#
If l2 <> 0 Then
    p = l1 / l2
    If p <> 0 Then FmtPercentOrBlank = Format$(p, "0%")
End If
End Function

'EHT=Custom
Private Function FmtTotal$(l&)
FmtTotal = Format$(l, "$#,##0")
End Function

'EHT=Custom
Private Function FmtTotalOrBlank$(l&)
If l <> 0 Then FmtTotalOrBlank = Format$(l, "$#,##0")
End Function

