VERSION 5.00
Begin VB.Form tabMailingList 
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
   Begin VB.PictureBox pctExportFrame 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3120
      Width           =   5415
      Begin VB.CommandButton btnExport 
         Caption         =   "Export to TB"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtExportFile 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label lblExport 
         Caption         =   "(entire email list + HC overrides with emails)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Timer tmrAnimation 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   4680
   End
   Begin VB.PictureBox pctPrintFrame 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5640
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3120
      Width           =   4695
      Begin VB.CommandButton btnScanForPaperSize 
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   9
         Top             =   285
         Width           =   615
      End
      Begin VB.TextBox txtPaperSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "Print Sel"
         Height          =   615
         Index           =   2
         Left            =   2760
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "Start at Sel"
         Height          =   615
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "Print All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "Note: Default printer must be changed before using scan or print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   11
         Top             =   630
         Width           =   4650
      End
   End
   Begin VB.PictureBox pctAnimation 
      Appearance      =   0  'Flat
      BackColor       =   &H0040FFFF&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3240
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   3735
      Begin VB.Label lblTitleAnimText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   30
         TabIndex        =   13
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   60
      End
   End
   Begin EJTSClients.CustomListbox lstSection 
      Height          =   1005
      Index           =   3
      Left            =   5880
      TabIndex        =   2
      Top             =   960
      Width           =   1005
      _ExtentX        =   0
      _ExtentY        =   0
      DisplayMode     =   0
      MultiSel        =   -1  'True
   End
   Begin EJTSClients.CustomListbox lstSection 
      Height          =   1005
      Index           =   2
      Left            =   4800
      TabIndex        =   1
      Top             =   960
      Width           =   1005
      _ExtentX        =   0
      _ExtentY        =   0
      DisplayMode     =   0
      MultiSel        =   -1  'True
   End
   Begin EJTSClients.CustomListbox lstSection 
      Height          =   1005
      Index           =   1
      Left            =   3720
      TabIndex        =   0
      Top             =   960
      Width           =   1005
      _ExtentX        =   0
      _ExtentY        =   0
      DisplayMode     =   0
      MultiSel        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "These listboxes must be Multi-Select!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblClientCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   10440
      TabIndex        =   22
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblCriteria 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   20
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblCriteria 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   19
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblCriteria 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   "WARNING: This tab should only be used on a new database file, after the proforma!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   840
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   10560
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "No Organizer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   15
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Email Organizer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Hard-copy Organizer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   17
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "tabMailingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabMailingList"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private LastClientIDSelected&(1 To 3)
Private LastScrollPosition&(1 To 3)
Public LastColumnFocused&

Public SkipChangeEvents As Boolean

Private animStartX!
Private animStartY!
Private animEndX!
Private animEndY!
Private animStepX!
Private animStepY!
Private Const animNumSteps& = 20
Private animStep&
Private NewItem_Lst&
Private NewItem_Index&

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Standard
Private Function ITab_CreateGDIObjects() As Boolean
On Error GoTo ERR_HANDLER


Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CreateGDIObjects", Err
End Function

'EHT=Standard
Private Function ITab_InitializeAfterDBLoad() As Boolean
On Error GoTo ERR_HANDLER

Dim a&
lblWarning.Visible = (FileToOpen_Year <> (Year(Date) - 1))
For a = 1 To 3
    lblTitle(a).Tag = lblTitle(a).Caption
    LastClientIDSelected(a) = -1
Next a

txtPaperSize.Enabled = ActiveDBInstance.IsWriteable
btnScanForPaperSize.Enabled = ActiveDBInstance.IsWriteable

txtExportFile.Text = DataFilesPath & "EJ Tax Service.csv"

Dim e As Boolean
e = lstSection(1).MultiSel
e = e And lstSection(2).MultiSel
e = e And lstSection(3).MultiSel
If Not e Then Err.Raise 1, , "lstSection not set to Multi-Select! Some features will not work properly!"

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

Dim a&, pft&
pft = DB_GetSetting(ActiveDBInstance, "Prep fee threshold - receive organizer")
lblCriteria(EmailOrganizer).Caption = "[Not dec'd, LY Completed, LY not IPTE, LYPrepFee > " & pft & "] and at least 1 email"
lblCriteria(HardCopyOrganizer).Caption = "[Not dec'd, LY Completed, LY not IPTE, LYPrepFee > " & pft & "] and no email address"
lblCriteria(NoOrganizer).Caption = "Everything else"
frmMain.ShowPopupInfo "Loading...", -1
DoEvents
For a = 1 To 3
    LastClientIDSelected(a) = lstSection(a).SelectedClientID
    LastScrollPosition(a) = lstSection(a).TopIndex
    lstSection(a).SetRedraw False
    lstSection(a).Clear
Next
For a = 0 To ActiveDBInstance.Clients_Count - 1
    AddToAppropriateList a, False
Next a
For a = 1 To 3
    lstSection(a).AddItem 0, -1
    lstSection(a).AddItem 2, -3
    lstSection(a).SetRedraw True
Next a
UpdateTotals
ReturnPrevSel
frmMain.HidePopupInfo

SkipChangeEvents = True
txtPaperSize.Text = DB_GetSetting(ActiveDBInstance, "_MailingList-PaperSize")
SkipChangeEvents = False

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr lstSection(LastColumnFocused)

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


Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_DestroyGDIObjects", Err
End Function

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

Dim l&, t&, w&, h&

w = (Me.ScaleWidth - 16) / 3
t = lblTitle(1).Height + lblCriteria(1).Height + 4
h = Me.ScaleHeight - t
If lblWarning.Visible Then
    lblWarning.Move 0, Me.ScaleHeight - lblWarning.Height, Me.ScaleWidth
    h = h - lblWarning.Height - 4
End If
lblClientCount.Move Me.ScaleWidth - lblClientCount.Width, 0

l = 0
lstSection(1).Move l, t, w, h - pctExportFrame.Height - 4
lblTitle(1).Move l, 0, w
lblCriteria(1).Move l, lblTitle(1).Height, w
pctExportFrame.Move l, t + lstSection(1).Height + 4, w
txtExportFile.Width = w
lblExport.Move btnExport.Width + 8, btnExport.Top + (btnExport.Height / 2) - (lblExport.Height / 2), pctExportFrame.ScaleWidth - btnExport.Width - 8

l = w + 8
lstSection(2).Move l, t, w, h - pctPrintFrame.Height - 4
lblTitle(2).Move l, 0, w
lblCriteria(2).Move l, lblTitle(2).Height, w
pctPrintFrame.Move l + (w / 2) - (pctPrintFrame.Width / 2), t + lstSection(2).Height + 4

l = (w * 2) + 16
lstSection(3).Move l, t, Me.ScaleWidth - l, h
lblTitle(3).Move l, 0, Me.ScaleWidth - l
lblCriteria(3).Move l, lblTitle(2).Height, w
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
Private Sub btnExport_Click()
On Error GoTo ERR_HANDLER

If Not btnExport.Enabled Then Exit Sub

SetFocusWithoutErr lstSection(MailingListStatus.HardCopyOrganizer)
ExportClients_ToThunderbird

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnExport_Click", Err
End Sub

'EHT=Standard
Private Sub btnPrint_Click(Index As Integer)
On Error GoTo ERR_HANDLER

If Not btnPrint(Index).Enabled Then Exit Sub

SetFocusWithoutErr lstSection(MailingListStatus.HardCopyOrganizer)
PrintClients Index

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnPrint_Click", Err
End Sub

'EHT=Standard
Private Sub btnScanForPaperSize_Click()
On Error GoTo ERR_HANDLER

If Not btnScanForPaperSize.Enabled Then Exit Sub

SetFocusWithoutErr lstSection(MailingListStatus.HardCopyOrganizer)
ScanForPaperSize

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnScanForPaperSize_Click", Err
End Sub

'EHT=Standard
Private Sub lstSection_DblClick(Index As Integer)
On Error GoTo ERR_HANDLER

Dim frm As frmClientEditPost, cID&
'Don't check .Enabled, because sometimes this code is called without showing the menu first

cID = lstSection(Index).SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item
Set frm = New frmClientEditPost
If frm.Form_Show(cID, fEdit) Then   'This will mark changed if necessary
    frmMain.DayTotal_Update
    lstSection(Index).Repaint
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSection_DblClick", Err
End Sub

'EHT=Standard
Private Sub lstSection_GotFocus(Index As Integer)
On Error GoTo ERR_HANDLER

LastColumnFocused = Index

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSection_GotFocus", Err
End Sub

'EHT=Standard
Private Sub lstSection_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

frmMain.IdleSetAction

Select Case KeyCode
Case vbKeyReturn
    KeyCode = 0
    Select Case Shift
    Case vbCtrlMask
        lstSection_DblClick (Index)
    End Select

Case vbKeyLeft
    KeyCode = 0
    If Index > 1 Then SetFocusWithoutErr lstSection(Index - 1)
Case vbKeyRight
    KeyCode = 0
    If Index < 3 Then SetFocusWithoutErr lstSection(Index + 1)

Case vbKeyA
    KeyCode = 0
    ProcessKey Index, MailingListStatus.Auto
Case vbKeyH
    KeyCode = 0
    ProcessKey Index, MailingListStatus.HardCopyOrganizer
Case vbKeyE
    KeyCode = 0
    ProcessKey Index, MailingListStatus.EmailOrganizer
Case vbKeyN
    KeyCode = 0
    ProcessKey Index, MailingListStatus.NoOrganizer
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSection_KeyDown", Err
End Sub

'EHT=Standard
Private Sub lstSection_KeyPressByCode(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyPress KeyCode: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event
frmMain.IdleSetAction

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSection_KeyPressByCode", Err
End Sub

'EHT=Standard
Private Sub lstSection_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSection_KeyUp", Err
End Sub

'EHT=Standard
Private Sub lstSection_TabToNextControl(Index As Integer, Reverse As Boolean)
On Error GoTo ERR_HANDLER

TabToNextControl Me, False, Reverse

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSection_TabToNextControl", Err
End Sub

'EHT=Standard
Private Sub txtPaperSize_Change()
On Error GoTo ERR_HANDLER

If Not txtPaperSize.Enabled Then Exit Sub
If Not SkipChangeEvents Then
    Dim p&
    p = Val(txtPaperSize.Text)
    DB_SetSetting ActiveDBInstance, "_MailingList-PaperSize", p, sLng
    txtPaperSize.Text = p
    frmMain.SetChangedFlagAndIndication
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtPaperSize_Change", Err
End Sub

'EHT=ResumeNext
Private Sub tmrAnimation_Timer()
On Error Resume Next

If animStep = animNumSteps Then
    pctAnimation.Visible = False
    tmrAnimation.Enabled = False
End If
animStartX = animStartX + animStepX
animStartY = animStartY + animStepY
pctAnimation.Move animStartX, animStartY
animStep = animStep + 1
End Sub

'EHT=Standard
Function AddToAppropriateList(cindex&, SelectNewItem As Boolean) As Long
On Error GoTo ERR_HANDLER

Dim li&, sec&
With ActiveDBInstance.Clients(cindex).c
    If .MailingListStatus = MailingListStatus.Auto Then
        sec = 3
        If ((Len(.Person1.First) > 0) And (.Person1.dod = NullLong)) Or _
           ((Len(.Person2.First) > 0) And (.Person2.dod = NullLong)) Then
            'At least one of the people is alive...
            If Flag_IsSet(.LastYear_Flags, CompletedReturn) And _
                (Not Flag_IsSet(.LastYear_Flags, IncPtnrTrustEstate)) And _
                (.LastYear_PrepFee > DB_GetSetting(ActiveDBInstance, "Prep fee threshold - receive organizer")) Then
                '... and certain conditions are met
                If (.Person1.Email = "") And (.Person2.Email = "") Then
                    'No email...
                    li = MailingListStatus.HardCopyOrganizer
                Else
                    'At least 1 email...
                    li = MailingListStatus.EmailOrganizer
                End If
            Else
                li = MailingListStatus.NoOrganizer
            End If
        Else
            'None of the people are alive...
            li = MailingListStatus.NoOrganizer
        End If
    Else
        li = .MailingListStatus
        sec = 1
    End If
End With
NewItem_Lst = li
NewItem_Index = lstSection(li).AddItem(sec, cindex)
If NewItem_Index <> LB_ERR Then
    If SelectNewItem Then lstSection(li).ListIndex = NewItem_Index
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "AddToAppropriateList", Err
End Function

'EHT=Cleanup1
Sub ExportClients_ToThunderbird()
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim fh As CMNMOD_CFileHandler, a&, cID&, cindex&, l$(36)
Set fh = OpenFile(txtExportFile.Text, mLineByLine_Output)
'             0          1         2            3        4             5               6           7          8          9          10           11            12           13             14        15         16           17           18           19             20        21         22           23           24        25         26           27         28         29         30          31        32       33       34       35       36
fh.WriteLine "First Name,Last Name,Display Name,Nickname,Primary Email,Secondary Email,Screen Name,Work Phone,Home Phone,Fax Number,Pager Number,Mobile Number,Home Address,Home Address 2,Home City,Home State,Home ZipCode,Home Country,Work Address,Work Address 2,Work City,Work State,Work ZipCode,Work Country,Job Title,Department,Organization,Web Page 1,Web Page 2,Birth Year,Birth Month,Birth Day,Custom 1,Custom 2,Custom 3,Custom 4,Notes"
l$(25) = ""
For a = 0 To lstSection(MailingListStatus.EmailOrganizer).ListCount - 1
    cID = lstSection(MailingListStatus.EmailOrganizer).ItemClientID(a)
    If cID >= 0 Then
        cindex = DB_FindClientIndex(ActiveDBInstance, cID)
        GoSub exp
    End If
Next a
l$(25) = "Hard Copy"
For a = 0 To lstSection(MailingListStatus.HardCopyOrganizer).ListCount - 1
    cID = lstSection(MailingListStatus.HardCopyOrganizer).ItemClientID(a)
    If cID >= 0 Then
        cindex = DB_FindClientIndex(ActiveDBInstance, cID)
        If ActiveDBInstance.Clients(cindex).c.MailingListStatus = MailingListStatus.HardCopyOrganizer Then
            If (ActiveDBInstance.Clients(cindex).c.Person1.Email <> "") Or (ActiveDBInstance.Clients(cindex).c.Person2.Email <> "") Then
                GoSub exp
            End If
        End If
    End If
Next a

ShowInfoMsg "Email list has been successfully exported to '" & txtExportFile.Text & "'."

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Sub
exp:
    With ActiveDBInstance.Clients(cindex).c
        l$(0) = FormatTextForCSV(FormatClientName(fExport_First, ActiveDBInstance.Clients(cindex).c))
        l$(1) = FormatTextForCSV(FormatClientName(fExport_Last, ActiveDBInstance.Clients(cindex).c))
        l$(2) = """" & FormatTextForCSV(FormatClientName(fExport_Display, ActiveDBInstance.Clients(cindex).c)) & """"
        l$(4) = LCase(FormatTextForCSV(.Person1.Email))
        l$(5) = LCase(FormatTextForCSV(.Person2.Email))
        l$(8) = FieldToString(.PhoneHome, mPhone)
        l$(7) = FieldToString(.Person1.Phone, mPhone)
        l$(11) = FieldToString(.Person2.Phone, mPhone)
        fh.WriteLine Join(l$, ",")
    End With
Return

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ExportClients_ToThunderbird", Err, INCLEANUP: Resume CLEANUP
End Sub

'EHT=Cleanup1
Sub PrintClients(i%)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

Dim s&, c&, a&, cID&
Dim ClientIDsToPrint() As Long, ClientIDsCount&

'######################################################################
'Note: Since it's a multi-sel listbox, we can't just ask for .ListIndex
'######################################################################

frmMain.ShowPopupInfo "Printing...", -1

'Values for i: 0:All, 1:FromSel, 2:OnlySel
c = lstSection(MailingListStatus.HardCopyOrganizer).ListCount
If i = 2 Then
    'PrintOnlySelected...
    For a = 0 To c - 1
        If lstSection(MailingListStatus.HardCopyOrganizer).Selected(a) Then
            cID = lstSection(MailingListStatus.HardCopyOrganizer).ItemClientID(a)
            If cID >= 0 Then
                ReDim Preserve ClientIDsToPrint(ClientIDsCount)
                ClientIDsToPrint(ClientIDsCount) = cID
                ClientIDsCount = ClientIDsCount + 1
            End If
        End If
    Next a
Else
    If i = 0 Then
        'PrintAll...
        s = 0
    ElseIf i = 1 Then
        'Print From Selection...
        s = -1
        For a = 0 To c - 1
            If lstSection(MailingListStatus.HardCopyOrganizer).Selected(a) Then
                s = a
                Exit For
            End If
        Next a
        If s < 0 Then
            GoTo CLEANUP
        End If
    End If
    For a = s To c - 1
        cID = lstSection(MailingListStatus.HardCopyOrganizer).ItemClientID(a)
        If cID >= 0 Then
            ReDim Preserve ClientIDsToPrint(ClientIDsCount)
            ClientIDsToPrint(ClientIDsCount) = cID
            ClientIDsCount = ClientIDsCount + 1
        End If
    Next a
End If

If ClientIDsCount > 0 Then
    Printer.PaperSize = DB_GetSetting(ActiveDBInstance, "_MailingList-PaperSize")
    Printer.ScaleMode = 1
    Printer.Font.Name = "Sans Serif 10cpi"
    Printer.Font.SIZE = 12
    For a = 0 To ClientIDsCount - 1
        c = DB_FindClientIndex(ActiveDBInstance, ClientIDsToPrint(a))
        If c < 0 Then
            ShowErrorMsg "Invalid client ID #" & ClientIDsToPrint(a)
            GoTo CLEANUP
        End If
        With ActiveDBInstance.Clients(c).c
            Printer.Print FormatClientName(fPrintLabels, ActiveDBInstance.Clients(c).c)
            Printer.CurrentY = Printer.CurrentY + 40
            Printer.Print UCase$(.AddressStreet)
            Printer.CurrentY = Printer.CurrentY + 40
            Printer.Print UCase$(.AddressCity) & ", " & UCase$(.AddressState) & " " & UCase$(.AddressZipCode)
            Printer.CurrentY = Printer.CurrentY + 40
            Printer.Print ""
            Printer.CurrentY = Printer.CurrentY + 40
            Printer.Print ""
            Printer.CurrentY = Printer.CurrentY + 40
            Printer.Print ""
            Printer.CurrentY = Printer.CurrentY + 40
        End With
    Next a
    Printer.EndDoc
End If

CLEANUP: INCLEANUP = True
    frmMain.HidePopupInfo

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "PrintClients", Err, INCLEANUP: Resume CLEANUP
End Sub

'EHT=Standard
Sub ProcessKey(Index As Integer, nli As MailingListStatus)
On Error GoTo ERR_HANDLER

If Not ActiveDBInstance.IsWriteable Then Exit Sub

Dim li&, cID&, cindex&, animStartPos&, animText$

li = lstSection(Index).ListIndex
If li = LB_ERR Then Exit Sub       'No selection
cID = lstSection(Index).ItemClientID(li)
If cID = LB_ERR Then Exit Sub      'Separator item, skip
cindex = DB_FindClientIndex(ActiveDBInstance, cID)
If cindex < 0 Then Exit Sub

'Remove item from list
With lstSection(Index)
    animStartPos = li - lstSection(Index).TopIndex
    .RemoveItem li
    If li >= .ListCount Then li = .ListCount - 1
    .ListIndex = li
End With

'Change ML flag
ActiveDBInstance.Clients(cindex).c.MailingListStatus = nli

'Add to new list
AddToAppropriateList cindex, True

UpdateTotals
frmMain.SetChangedFlagAndIndication

DoEvents
StartAnimation CLng(Index), animStartPos, NewItem_Lst, NewItem_Index - lstSection(NewItem_Lst).TopIndex, _
        animText$ = FormatClientName(fMailingList, ActiveDBInstance.Clients(cindex).c)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ProcessKey", Err
End Sub

'EHT=Standard
Sub ReturnPrevSel()
On Error GoTo ERR_HANDLER

Dim a&, b&
For a = 1 To 3
    With lstSection(a)
        If .ListIndex < 0 Then
            If LastClientIDSelected(a) < 0 Then
                If .ListCount > 0 Then .ListIndex = 0
            Else
                For b = 0 To .ListCount - 1
                    If .ItemClientID(b) = LastClientIDSelected(a) Then
                        .ListIndex = b
                        .TopIndex = LastScrollPosition(a)
                        Exit For
                    End If
                Next b
            End If
        End If
    End With
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ReturnPrevSel", Err
End Sub

'EHT=Cleanup1
Sub ScanForPaperSize()
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

If Not ActiveDBInstance.IsWriteable Then Exit Sub

'Search for paper size of 8.5 x 12
Dim p&, ps&
Dim fh As CMNMOD_CFileHandler
Set fh = OpenFile(AppPath & "EJTSClients-PaperSizeScan.log", mLineByLine_Output)
SkipChangeEvents = True
ps = 0
For p = 1 To 512
    txtPaperSize.Text = p: DoEvents
    On Error Resume Next
    Printer.PaperSize = p
    On Error GoTo ERR_HANDLER
    If Printer.PaperSize = p Then fh.WriteLine Printer.PaperSize & vbTab & Round(Printer.Height / 1440, 2) & " x " & Round(Printer.Width / 1440, 2)
    If Printer.Height = 12 * 1440 Then
        If Printer.Width = 8.5 * 1440 Then
            If ps = 0 Then ps = p
        End If
    End If
Next p
If ps = 0 Then
    txtPaperSize.Text = DB_GetSetting(ActiveDBInstance, "_MailingList-PaperSize")
    ShowErrorMsg "Unable to find new paper size code! Textbox reverted to previous one."
Else
    DB_SetSetting ActiveDBInstance, "_MailingList-PaperSize", ps, sLng
    txtPaperSize.Text = ps
    frmMain.SetChangedFlagAndIndication
End If

CLEANUP: INCLEANUP = True
    SkipChangeEvents = False
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ScanForPaperSize", Err, INCLEANUP: Resume CLEANUP
End Sub

'EHT=Standard
Sub StartAnimation(startLst&, startPos&, endLst&, endPos&, animText$)
On Error GoTo ERR_HANDLER

animStartX = lstSection(startLst).Left
animStartY = lstSection(startLst).Top + (startPos * 18)
animEndX = lstSection(endLst).Left
animEndY = lstSection(endLst).Top + (endPos * 18)
animStepX = (animEndX - animStartX) / animNumSteps
animStepY = (animEndY - animStartY) / animNumSteps
animStep = 0
pctAnimation.Move animStartX, animStartY, lstSection(startLst).Width
lblTitleAnimText.Caption = animText$
pctAnimation.Visible = True
tmrAnimation.Enabled = True

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "StartAnimation", Err
End Sub

'EHT=Standard
Sub UpdateTotals()
On Error GoTo ERR_HANDLER

Dim a&, b&, c&, cID&, liOverrides&, totalOverrides&, liAuto&, totalAuto&, secnum&, totalNoOrgCO&
For a = 1 To 3
    lblTitle(a).Caption = lblTitle(a).Tag & " (" & (lstSection(a).ListCount - 2) & ")"

    liOverrides = -1
    totalOverrides = 0
    liAuto = -1
    totalAuto = 0
    secnum = -1
    c = lstSection(a).ListCount
    For b = 0 To c - 1
        cID = lstSection(a).ItemClientID(b)
        Select Case cID
        Case -1
            liOverrides = b
            secnum = cID
        Case -3
            liAuto = b
            secnum = cID
        Case Else
            Select Case secnum
            Case -1
                totalOverrides = totalOverrides + 1
            Case -3
                totalAuto = totalAuto + 1
                If a = NoOrganizer Then
                    With ActiveDBInstance.Clients(DB_FindClientIndex(ActiveDBInstance, cID)).c
                        If Flag_IsSet(.LastYear_Flags, CompletedReturn) Then totalNoOrgCO = totalNoOrgCO + 1
                    End With
                End If
            End Select
        End Select
    Next b
    DoEvents
    lstSection(a).RemoveItem liOverrides
    lstSection(a).AddItem 0, -1, "Overrides (" & totalOverrides & ") - Press '" & Mid$("EHN", a, 1) & "'"
    lstSection(a).RemoveItem liAuto
    If a < NoOrganizer Then
        lstSection(a).AddItem 2, -3, "Auto (" & totalAuto & ") - Press 'A'"
    Else
        'The NoOrganizer list, Auto subsection is subtotaled for completed & not-completed returns
        lstSection(a).AddItem 2, -3, "Auto (" & totalNoOrgCO & " co + " & (totalAuto - totalNoOrgCO) & " nc = " & totalAuto & ") - Press 'A'"
    End If
Next a
lblClientCount.Caption = "Total:" & vbCrLf & ActiveDBInstance.Clients_Count

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UpdateTotals", Err
End Sub

