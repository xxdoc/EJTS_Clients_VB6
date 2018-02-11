VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "EJTS Clients"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1337
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctSecondFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -480
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox pctInitialFocus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -840
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox pctPopupInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   4200
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
      Begin VB.Timer tmrPopupInfo 
         Enabled         =   0   'False
         Left            =   1560
         Top             =   1320
      End
      Begin VB.Shape shpPopupInfo 
         BorderWidth     =   3
         Height          =   765
         Left            =   15
         Top             =   15
         Width           =   465
      End
      Begin VB.Label lblPopupInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3585
         TabIndex        =   12
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   165
      End
   End
   Begin VB.Timer tmrAutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3600
      Top             =   1080
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save DB"
      Height          =   615
      Left            =   18360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btnNewClient 
      Caption         =   "New Client..."
      Height          =   615
      Left            =   18120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   1815
   End
   Begin VB.Timer tmrDate 
      Interval        =   1000
      Left            =   3600
      Top             =   120
   End
   Begin VB.ListBox CHOS_lstClients 
      Height          =   1215
      IntegralHeight  =   0   'False
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ComboBox SRCH_cboSpecialSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3615
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Width           =   19695
      _ExtentX        =   34740
      _ExtentY        =   1085
      MultiRow        =   -1  'True
      TabFixedHeight  =   661
      HotTracking     =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Schedule "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Search "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Pull Files "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Sat Check "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Extra Charges "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Bookkeeping "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Unpaid "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Stats "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Mailing Lists "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Log "
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Prefs "
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label DTOT_lblDayTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   9840
      TabIndex        =   22
      ToolTipText     =   "Click to recalculate"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label DTOT_lblDayTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   9840
      TabIndex        =   21
      ToolTipText     =   "Click to recalculate"
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label DTOT_lblDayTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   9840
      TabIndex        =   20
      ToolTipText     =   "Click to recalculate"
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label DTOT_lblDayTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   9840
      TabIndex        =   19
      ToolTipText     =   "Click to recalculate"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label DTOT_lblDayTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Appts Made Today: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   14400
      TabIndex        =   15
      ToolTipText     =   "Click to recalculate"
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label SRCH_lblSpecialSearchEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "(edit list)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   16560
      MouseIcon       =   "frmMain.frx":57E2
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblReadOnlyMode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Read-only Mode)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label SRCH_lblSpecialSearch 
      BackStyle       =   0  'Transparent
      Caption         =   "Favorite Searches:"
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
      Left            =   14400
      TabIndex        =   14
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label DTOT_lblDayTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Daily Total: $0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   14400
      TabIndex        =   13
      ToolTipText     =   "Click to recalculate"
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label CHOS_lblApptInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total: 123 QA (change with +/-)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10:43 AM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "February 21, 2008"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblDayOfWeek 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label CHOS_lblClients 
      BackStyle       =   0  'Transparent
      Caption         =   "Chosen clients:"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "frmMain"

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Public Enum enumTabName
    vSchedule
    vSearch
    vPullFiles
    vSatCheck
    vExtraCharges
    vBookkeeping
    vUnpaid
    vStatistics
    vMailingList
    vLogFile
    vSettings
End Enum
Public CurTab As enumTabName
Private Tabs(0 To 10) As ITab

Public CHOS_NumMinutes&
Public CHOS_NumSlots&
Public CHOS_NumSlotsBeforeOverride&
Public CHOS_NumSlots_Overridden As Boolean

'Global variables
Public PopupInfoActive As Boolean
Public DontCallChangeCurTab As Boolean

Private IdleNextTimeout As Date

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Function Form_Show() As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

Dim a&
If DEBUGMODE Then
    Me.Caption = "EJTS Clients vXXX.XXX.XXX"
    Me.Icon = LoadPicture(AppPath & "DebugMode.ico")
    Me.BackColor = &HC0E0FF
Else
    Me.Caption = "EJTS Clients v" & App.Major & "." & App.Minor & "." & App.Revision
End If
Me.Tag = Me.Caption
btnSave.Tag = btnSave.Caption

'The 'New' commands are important here!
Set tabSchedule = New tabSchedule:          Set Tabs(vSchedule) = tabSchedule
Set tabSearch = New tabSearch:              Set Tabs(vSearch) = tabSearch
Set tabPullFiles = New tabPullFiles:        Set Tabs(vPullFiles) = tabPullFiles
Set tabSatCheck = New tabSatCheck:          Set Tabs(vSatCheck) = tabSatCheck
Set tabExtraCharges = New tabExtraCharges:  Set Tabs(vExtraCharges) = tabExtraCharges
Set tabBookkeeping = New tabBookkeeping:    Set Tabs(vBookkeeping) = tabBookkeeping
Set tabUnpaid = New tabUnpaid:              Set Tabs(vUnpaid) = tabUnpaid
Set tabStatistics = New tabStatistics:      Set Tabs(vStatistics) = tabStatistics
Set tabMailingList = New tabMailingList:    Set Tabs(vMailingList) = tabMailingList
Set tabLogFile = New tabLogFile:            Set Tabs(vLogFile) = tabLogFile
Set tabSettings = New tabSettings:          Set Tabs(vSettings) = tabSettings

'Position all the sub-forms
Dim f As Form
For a = 0 To UBound(Tabs)
    If Not Tabs(a) Is Nothing Then
        Set f = Tabs(a)
        SetParent f.hwnd, Me.hwnd
        Dim lllll&
        lllll = GetWindowLong(f.hwnd, GWL_STYLE)
        lllll = Not ((Not (lllll Or WS_CHILD)) Or WS_BORDER)
        SetWindowLong f.hwnd, GWL_STYLE, lllll
    End If
Next a

'Initiate modules
For a = 0 To UBound(Tabs)
    Tabs(a).CreateGDIObjects
Next a

'############ Initialize frmMain
SetTabStops CHOS_lstClients.hwnd, 20, 40
tabSearch.ClearAll
tmrDate_Timer
CHOS_CalculateTotal

Me.Show

'############ Load database
ShowPopupInfo "Loading Database", -1
Me.Caption = FileToOpen_Year & " Tax Season" & " - " & DataFilesPath & " - " & Me.Tag
Me.Tag = Me.Caption
If Not DB_Load(DataFilesPath & "EJTSClients" & FileToOpen_Year & ".dat", ActiveDBInstance) Then GoTo CLEANUP
ActiveDBInstance.IsWriteable = Not FileToOpen_OpenReadOnly

#If False Then
    Dim sd As Date, ed As Date, b&
    sd = ActiveDBInstance.ApptBitmap_StartDate
    ed = ActiveDBInstance.ApptBitmap_StartDate + ActiveDBInstance.ApptBitmap_Count - 1
    Dim nsd As Date, ned As Date, napptbm() As Long, nc&
    nsd = DateSerial(Year(sd), 1, 1)
    ned = DateSerial(Year(sd), 12, 31)
    nc = ned - nsd + 1
    ReDim napptbm(nc - 1, Appointment_NumSlotsUB)
    For a = 0 To nc - 1
        If ((nsd + a) < sd) Or ((nsd + a) > ed) Then
            Debug.Print "" & a & " <- NEW"
            For b = 0 To Appointment_NumSlotsUB
                napptbm(a, b) = -99999999
            Next b
        Else
            Debug.Print "" & a & " <- " & (nsd + a - sd)
            For b = 0 To Appointment_NumSlotsUB
                napptbm(a, b) = ActiveDBInstance.ApptBitmap(nsd + a - sd, b)
            Next b
        End If
    Next a
    ActiveDBInstance.ApptBitmap = napptbm
    ActiveDBInstance.ApptBitmap_Count = nc
    ActiveDBInstance.ApptBitmap_StartDate = nsd
#End If

'Read a few settings that otherwise wouldn't get initialized right away
DB_GetSetting ActiveDBInstance, SETTING_SLOTBREAKPOINT

ClearChangedIndication
DayTotal_Update
HidePopupInfo

'Initialize our form
lblReadOnlyMode.Visible = Not ActiveDBInstance.IsWriteable
btnNewClient.Enabled = ActiveDBInstance.IsWriteable
btnSave.Enabled = ActiveDBInstance.IsWriteable
SRCH_lblSpecialSearchEdit.Visible = ActiveDBInstance.IsWriteable

'Initialize the tabs
For a = 0 To UBound(Tabs)
    If Not Tabs(a) Is Nothing Then
        Tabs(a).InitializeAfterDBLoad
    End If
Next a

'The -1 causes the first call to ShowDate (below) to not update the schedule
'  (we want the next line ChangeCurTab to do it)
CurTab = -1
tabSchedule.ShowDate Date
ChangeCurTab vSchedule, False

'If this is a new month, and the user forgot to create a snapshot, create one automatically
tabStatistics.CreateAutoSnapshotIfNewMonth

Form_Show = True

CLEANUP: INCLEANUP = True
    If HASERROR Then
        ActiveDBInstance.Loaded = False
        HidePopupInfo
        Unload Me
    End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Function

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

With pctPopupInfo
    .Left = (Me.ScaleWidth / 2) - (.Width / 2)
    .Top = (Me.ScaleHeight / 2) - (.Height / 2)
End With
With TabStrip
    .Width = Me.ScaleWidth - .Left - 8
    .Height = Me.ScaleHeight - .Top - 4

    Dim f As Form
    Set f = Tabs(CurTab)
    f.Move (.ClientLeft + 8) * Screen.TwipsPerPixelX, (.ClientTop + 8) * Screen.TwipsPerPixelY, (.ClientWidth - 16) * Screen.TwipsPerPixelX, (.ClientHeight - 16) * Screen.TwipsPerPixelY
End With
End Sub

'EHT=Standard
Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

IdleSetAction

Select Case KeyCode
Case vbKeyEscape
    KeyCode = 0
    If pctPopupInfo.Visible Then HidePopupInfo
    If tabSchedule.tmrFlashAppt.Enabled Then tabSchedule.StopFlashAppt
    If tabSchedule.ScheduleMode = sReschedule Then
        tabSchedule.ChangeScheduleMode sView
        tabSchedule.DrawSchedule
    Else
        If (CurTab = vSearch) And ((tabSearch.txtSearch.Text <> "") Or (tabSearch.lstResults.ListCount > 0)) Then
            tabSearch.ClearAll
        Else
            tabSearch.ClearAll
            If CHOS_lstClients.ListCount > 0 Then
                CHOS_Clear
                tabSchedule.ChangeScheduleMode sView
            End If
            ChangeCurTab vSchedule, False
            tabSchedule.ShowDate Date
        End If
    End If
Case vbKeyPageUp
    If Shift = vbCtrlMask Then
        If CurTab = 0 Then
            ChangeCurTab (TabStrip.Tabs.Count - 1), False
        Else
            ChangeCurTab CurTab - 1, False
        End If
        KeyCode = 0
    End If
Case vbKeyPageDown
    If Shift = vbCtrlMask Then
        If CurTab = (TabStrip.Tabs.Count - 1) Then
            ChangeCurTab 0, False
        Else
            ChangeCurTab CurTab + 1, False
        End If
        KeyCode = 0
    End If
Case vbKeyTab
    If Shift = (vbCtrlMask Or vbShiftMask) Then
        If CurTab = 0 Then
            ChangeCurTab (TabStrip.Tabs.Count - 1), False
        Else
            ChangeCurTab CurTab - 1, False
        End If
        KeyCode = 0
    ElseIf Shift = vbCtrlMask Then
        If CurTab = (TabStrip.Tabs.Count - 1) Then
            ChangeCurTab 0, False
        Else
            ChangeCurTab CurTab + 1, False
        End If
        KeyCode = 0
    End If
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=Standard
Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

'These get called from the tab-forms. Even if we're not using them yet, leave them here.

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyUp", Err
End Sub

'EHT=Standard
Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_HANDLER

'These get called from the tab-forms. Even if we're not using them yet, leave them here.

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyPress", Err
End Sub

'EHT=Cleanup2
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

Dim a&

tmrDate.Enabled = False
tmrAutoSave.Enabled = False

'Save database
If ActiveDBInstance.Loaded And ActiveDBInstance.Changed Then
    For a = 0 To UBound(Tabs)
        If Not Tabs(a) Is Nothing Then
            Tabs(a).SaveSettingsToDBBeforeClose
        End If
    Next a
    If DB_Save(ActiveDBInstance) Then
        ClearChangedIndication
        tabLogFile.WriteLine "Save"
    Else
        Cancel = True
        Exit Sub
    End If
End If

For a = 0 To UBound(Tabs)
    If Not Tabs(a) Is Nothing Then
        Tabs(a).DestroyGDIObjects
        Unload Tabs(a)
    End If
Next a

CLEANUP: INCLEANUP = True
    If HASERROR Then Cancel = True

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Unload", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Sub

'EHT=Standard
Private Sub pctInitialFocus_GotFocus()
On Error GoTo ERR_HANDLER

'If the entire program loses focus and then regains it, pctInitialFocus will get
'   the focus, since it's TabIndex=0. If the focus is passed back to the sub-form,
'   nothing happens. But if focus is changed to a fellow-control, and then passed
'   to the sub-form, it works for some reason.
Dim f As Form
Set f = Tabs(CurTab)
SetFocusWithoutErr pctSecondFocus
SetFocusWithoutErr f

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "pctInitialFocus_GotFocus", Err
End Sub

'EHT=ResumeNext
Public Sub tmrDate_Timer()
On Error Resume Next

Dim n As Date, nt As Date, cp As POINTAPI

n = Date
nt = Time

'Update the labels
lblDayOfWeek.Caption = Format$(n, "dddd")
lblDate.Caption = Format$(n, "mmmm d, yyyy")
lblTime.Caption = Format$(nt, "h:mm AM/PM")

'Move the red arrow
If CurTab = vSchedule Then tabSchedule.MoveRedArrow nt

'Idle code
GetCursorPos cp
'If (cp.X <> Idle_LastCursorPos.X) Or (cp.Y <> Idle_LastCursorPos.Y) Then
'    Idle_SetAction
'    Idle_LastCursorPos = cp
'End If
If IdleNextTimeout <> 0 Then
    If (n + nt) >= IdleNextTimeout Then
        IdleNextTimeout = 0
        If pctPopupInfo.Visible Then HidePopupInfo
        If tabSchedule.tmrFlashAppt.Enabled Then tabSchedule.StopFlashAppt
        If CHOS_lstClients.ListCount > 0 Then
            CHOS_Clear
            tabSchedule.ChangeScheduleMode sView
        End If
        ChangeCurTab vSchedule, False
        tabSchedule.ShowDate Date
    End If
End If
End Sub

'EHT=Standard
Private Sub btnNewClient_Click()
On Error GoTo ERR_HANDLER

If Not btnNewClient.Enabled Then Exit Sub

Dim frm As frmClientEditPost, nID&
SetFocusWithoutErr pctInitialFocus
Set frm = New frmClientEditPost
'Give it anything that has been typed, in case it is useful
If frm.Form_Show(nID, fNew, , tabSearch.txtSearch.Text) Then   'This will mark changed if necessary
    CHOS_Add nID, False
    'If Not SearchEditMode Then SRCH_Do False     'Redo search to make new client show up
    ChangeCurTab vSchedule, False
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnNewClient_Click", Err
End Sub

'EHT=Standard
Private Sub btnSave_Click()
On Error GoTo ERR_HANDLER

If Not btnSave.Enabled Then Exit Sub

Dim a&

SetFocusWithoutErr pctInitialFocus
ShowPopupInfo "Saving Database", -1
For a = 0 To UBound(Tabs)
    If Not Tabs(a) Is Nothing Then
        Tabs(a).SaveSettingsToDBBeforeClose
    End If
Next a
If DB_Save(ActiveDBInstance) Then
    ClearChangedIndication
    tabLogFile.WriteLine "Save"
End If
HidePopupInfo

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSave_Click", Err
End Sub

'EHT=Standard
Private Sub TabStrip_Click()
On Error GoTo ERR_HANDLER

Dim t As MSComctlLib.Tab
Static lt As MSComctlLib.Tab

Set t = TabStrip.SelectedItem
If Not DontCallChangeCurTab Then ChangeCurTab t.Index - 1, True
If Not lt Is Nothing Then lt.HighLighted = False
t.HighLighted = True
Set lt = t

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "TabStrip_Click", Err
End Sub

'EHT=ResumeNext
Private Sub tmrAutoSave_Timer()
On Error Resume Next

Dim a&
tmrAutoSave.Enabled = False
If ActiveDBInstance.Changed Then
    ShowPopupInfo "Auto-Save", -1
    For a = 0 To UBound(Tabs)
        If Not Tabs(a) Is Nothing Then
            Tabs(a).SaveSettingsToDBBeforeClose
        End If
    Next a
    If DB_Save(ActiveDBInstance) Then
        ClearChangedIndication
        tabLogFile.WriteLine "Auto-save"
    End If
    HidePopupInfo
End If
End Sub

'EHT=ResumeNext
Private Sub tmrPopupInfo_Timer()
On Error Resume Next

HidePopupInfo
End Sub

'EHT=Standard
Private Sub CHOS_lstClients_DblClick()
On Error GoTo ERR_HANDLER

Dim i&

SetFocusWithoutErr pctInitialFocus

'Remove selected client from chosen list
i = CHOS_lstClients.ListIndex
If i >= 0 Then CHOS_Remove i

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_lstClients_DblClick", Err
End Sub

'EHT=Standard
Private Sub CHOS_lstClients_GotFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr pctInitialFocus

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_lstClients_GotFocus", Err
End Sub

'EHT=Standard
Private Sub SRCH_cboSpecialSearch_Click()
On Error GoTo ERR_HANDLER

If tabSearch.SkipChangeEvents Then Exit Sub

Dim a&
a = SRCH_cboSpecialSearch.ListIndex
If a < 0 Then Exit Sub
ChangeCurTab vSearch, False
tabSearch.SkipChangeEvents = True
tabSearch.txtSearch.Text = ActiveDBInstance.SpecialSearches(a).SearchString
SRCH_cboSpecialSearch.Tag = ActiveDBInstance.SpecialSearches(a).DisplayName
tabSearch.SkipChangeEvents = False
SetFocusWithoutErr tabSearch.txtSearch  'Prevents SRCH_lstResults flicker
tabSearch.DoSearch
tabSearch.UpdateTabAsterisk
SetFocusWithoutErr pctInitialFocus

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SRCH_cboSpecialSearch_Click", Err
End Sub

'EHT=Standard
Private Sub SRCH_lblSpecialSearchEdit_Click()
On Error GoTo ERR_HANDLER

Dim frm As frmEditSearches
Set frm = New frmEditSearches
frm.Form_Show

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SRCH_lblSpecialSearchEdit_Click", Err
End Sub

'EHT=Standard
Private Sub DTOT_lblDayTotal_Click(Index As Integer)
On Error GoTo ERR_HANDLER

DayTotal_Update

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DTOT_lblDayTotal_Click", Err
End Sub

'EHT=Standard
Sub ChangeCurTab(ct As enumTabName, FromTabStripEvent As Boolean)
On Error GoTo ERR_HANDLER

Dim a&, oct As enumTabName, f As Form

oct = CurTab
CurTab = ct

If Not FromTabStripEvent Then
    Dim t As MSComctlLib.Tab
    Set t = TabStrip.Tabs(ct + 1)
    DontCallChangeCurTab = True
    t.Selected = True
    DontCallChangeCurTab = False
End If

For a = 0 To UBound(Tabs)
    Set f = Tabs(a)
    If a = ct Then
        If Not f.Visible Then
            'Use SetWindowPos to show without causing a Resize event like ShowWindow does
            SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
            f.Move (TabStrip.ClientLeft + 8) * Screen.TwipsPerPixelX, (TabStrip.ClientTop + 8) * Screen.TwipsPerPixelY, (TabStrip.ClientWidth - 16) * Screen.TwipsPerPixelX, (TabStrip.ClientHeight - 16) * Screen.TwipsPerPixelY
'            SetWindowPos f.hwnd, 0, TabStrip.ClientLeft + 8, TabStrip.ClientTop + 8, TabStrip.ClientWidth - 16, TabStrip.ClientHeight - 16, SWP_SHOWWINDOW
        End If
    Else
        If f.Visible Then
            SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        End If
    End If
Next a

SetFocusWithoutErr pctInitialFocus
Tabs(ct).AfterTabShown
If FromTabStripEvent Then
    Tabs(ct).SetDefaultFocus
Else
    'If called from code, no need to change focus of anything if the CurTab hasn't changed
    If ct <> oct Then
        Tabs(ct).SetDefaultFocus
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ChangeCurTab", Err
End Sub

'EHT=Standard
Sub SetChangedFlagAndIndication()
On Error GoTo ERR_HANDLER

If Not ActiveDBInstance.IsWriteable Then
    'This should never actually happen, since there are other protections in the code
    Err.Raise 1, , "Database has been opened in read-only mode, yet SetChangedFlagAndIndication has been called! Your changes will not actually be saved."
End If

ActiveDBInstance.Changed = True
Me.Caption = Me.Tag & " - CHANGED"
btnSave.Caption = btnSave.Tag & " (*)"
tmrAutoSave.Enabled = False
tmrAutoSave.Enabled = True

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetChangedFlagAndIndication", Err
End Sub

'EHT=Standard
Sub ClearChangedIndication()
On Error GoTo ERR_HANDLER

Me.Caption = Me.Tag
btnSave.Caption = btnSave.Tag

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ClearChangedIndication", Err
End Sub

'EHT=Standard
Sub ShowPopupInfo(i$, secondstoshow#)
On Error GoTo ERR_HANDLER

With lblPopupInfo
    .Caption = i$
    .Move 8, 8
End With
With pctPopupInfo
    .Height = lblPopupInfo.Height + 16
    .Width = lblPopupInfo.Width + 16
    .Left = (Me.ScaleWidth / 2) - (.Width / 2)
    .Top = (Me.ScaleHeight / 2) - (.Height / 2)
    .Visible = True
    .ZOrder 0
End With
shpPopupInfo.Move 1, 1, pctPopupInfo.ScaleWidth - 2, pctPopupInfo.ScaleHeight - 2
If secondstoshow > 0 Then
    tmrPopupInfo.Interval = secondstoshow * 1000
    tmrPopupInfo.Enabled = False
    tmrPopupInfo.Enabled = True
End If
PopupInfoActive = True

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ShowPopupInfo", Err
End Sub

'EHT=Standard
Sub HidePopupInfo()
On Error GoTo ERR_HANDLER

If Not PopupInfoActive Then Exit Sub
PopupInfoActive = False
tmrPopupInfo.Enabled = False
pctPopupInfo.Visible = False

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "HidePopupInfo", Err
End Sub

'EHT=None
Sub IdleSetAction()
IdleNextTimeout = Now + (1 / 24 / 60 * 5)
End Sub

'EHT=None
Sub IdlePauseTimeout()
IdleNextTimeout = 0
End Sub

'EHT=Standard
Sub DayTotal_Update()
On Error GoTo ERR_HANDLER

Dim a&, b&, cd&, ts&, tot&, numappt&, t$
Dim totalslots&, usedslots&, usedclients&, availslots&, lid&, st&

t$ = "*" & Format$(Date, "yyyy-mm-dd") & "????????Scheduled?appt:*"
cd = CLng(Date)
For a = 0 To ActiveDBInstance.Clients_Count - 1
    'Daily total
    If ActiveDBInstance.Clients(a).c.CompletionDate = cd Then
        If ActiveDBInstance.Clients(a).c.PrepFee <> NullLong Then tot = tot + ActiveDBInstance.Clients(a).c.PrepFee
    End If
    'Num appts
    If ActiveDBInstance.Clients(a).c.OpNotes Like t$ Then numappt = numappt + 1
Next a
For a = 0 To ActiveDBInstance.ExtraCharges_Count - 1
    If ActiveDBInstance.ExtraCharges(a).CompletionDate = cd Then
        If ActiveDBInstance.ExtraCharges(a).PrepFee <> NullLong Then tot = tot + ActiveDBInstance.ExtraCharges(a).PrepFee
    End If
Next a
DTOT_lblDayTotal(0).Caption = "Daily Total: " & FieldToString(tot, mDollar)
DTOT_lblDayTotal(1).Caption = "Appts Made Today: " & FieldToString(numappt, mNumber)

numappt = 0
lid = -1
st = CLng(Date) - ActiveDBInstance.ApptBitmap_StartDate
For cd = (ActiveDBInstance.ScheduleTemplateBreakpoint1 - ActiveDBInstance.ApptBitmap_StartDate) To (ActiveDBInstance.ScheduleTemplateBreakpoint2 - ActiveDBInstance.ApptBitmap_StartDate - 1)
    For ts = 0 To Appointment_NumSlotsUB
        a = DB_GetIDAtSlot(ActiveDBInstance, ActiveDBInstance.ApptBitmap_StartDate + cd, ts)
        If a >= 0 Then
            totalslots = totalslots + 1
            usedslots = usedslots + 1
            If a <> lid Then
                usedclients = usedclients + 1
                lid = a
            End If
        ElseIf a = Slot_Available Then
            totalslots = totalslots + 1
            If cd > st Then availslots = availslots + 1
        ElseIf a = Slot_MealBreak Or a = Slot_Reserved Then
        End If
    Next ts
Next cd
DTOT_lblDayTotal(2).Caption = "Season " & FieldToString(ActiveDBInstance.ScheduleTemplateBreakpoint1, mDateAsLong) & " - " & FieldToString(ActiveDBInstance.ScheduleTemplateBreakpoint2 - 1, mDateAsLong)
DTOT_lblDayTotal(3).Caption = "Total slots for season: " & FieldToString(totalslots, mNumber)
DTOT_lblDayTotal(4).Caption = FieldToString(usedslots, mNumber) & " slots used by " & FieldToString(usedclients, mNumber) & " appointments"
DTOT_lblDayTotal(5).Caption = FieldToString(availslots, mNumber) & " slots available after today"

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DayTotal_Update", Err
End Sub

'EHT=Standard
Sub CHOS_Add(ByVal cID&, GotoScheduleIfAlreadyChosen As Boolean)
On Error GoTo ERR_HANDLER

Dim i&, cindex&, foundindex&

'Find client
cindex = DB_FindClientIndex(ActiveDBInstance, cID)
If cindex < 0 Then Err.Raise 1, , "Client not found"

If DEBUGMODE Then
    'If it's null, then set it to -1 to signify we're creating a new appt
    If ApptBeingRescheduled.ID = NullLong Then
        ApptBeingRescheduled.ID = -1
        CHOS_lstClients.Visible = True
    End If
End If

'Check if client already chosen
foundindex = -1
For i = 0 To CHOS_lstClients.ListCount - 1
    If CHOS_lstClients.ItemData(i) = cID Then
        foundindex = i
        Exit For
    End If
Next i

If foundindex < 0 Then
    'Not found
    'Add to chosen list
    CHOS_Add2 cID, cindex

    With ActiveDBInstance.Clients(cindex).c
        If Flag_IsSet(.LastYear_Flags, NoNeedToFile) Then
            ShowInfoMsg FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c) & " was marked NNTF last year." & vbCrLf & vbCrLf & "Fill out No Need To File checklist."
        ElseIf (.LastYear_MinutesToComplete = NullLong) And (Not Flag_IsSet(.Flags, NewClient)) Then
            ShowInfoMsg "A return for " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c) & " was not completed last year (LYMin=NULL)." & vbCrLf & vbCrLf & "Fill out New Client checklist."
        End If
    End With
Else
    'Found
    If GotoScheduleIfAlreadyChosen Then
        tabSchedule.ShowDate Date
        frmMain.ChangeCurTab vSchedule, False
        Exit Sub
    Else
        CHOS_lstClients.RemoveItem i
    End If
End If

If Not DEBUGMODE Then
    CHOS_CalculateTotal
    tabSearch.lstResults.Repaint
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_Add", Err
End Sub

'EHT=Standard
Public Sub CHOS_Add2(cID&, cindex&)
On Error GoTo ERR_HANDLER

With ActiveDBInstance.Clients(cindex).c
    CHOS_lstClients.AddItem FieldToString(.LastYear_MinutesToComplete, mNumberOrNULL) & vbTab & FormatNumApptSlots(.NumApptSlotsToUse) & vbTab & FormatClientName(fChosenClients, ActiveDBInstance.Clients(cindex).c)
    CHOS_lstClients.ItemData(CHOS_lstClients.NewIndex) = cID
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_Add2", Err
End Sub

'EHT=Standard
Public Sub CHOS_CalculateTotal()
On Error GoTo ERR_HANDLER
Dim cclc&, i&, cindex&, totalminutes&, minforslots&
cclc = CHOS_lstClients.ListCount
If cclc = 0 Then
    CHOS_lblApptInfo.Caption = ""
    tabSchedule.ChangeScheduleMode sView
Else
    For i = 0 To cclc - 1
        cindex = DB_FindClientIndex(ActiveDBInstance, CHOS_lstClients.ItemData(i))
        With ActiveDBInstance.Clients(cindex).c
            If .NumApptSlotsToUse = 0 Then
                'If client has no override set for minutes...
                If .LastYear_MinutesToComplete = NullLong Then
                    'If no LY history and no NumSlots override set, then just assume it's a DA, like the new clients
                    minforslots = minforslots + (2 * DB_GetSetting(ActiveDBInstance, SETTING_SLOTBREAKPOINT))
                Else
                    'LY has history, so use it...
                    minforslots = minforslots + .LastYear_MinutesToComplete
                    If Flag_IsSet(.LastYear_Flags, DroppedOff) Or Flag_IsSet(.LastYear_Flags, MailedIn) Then
                        'DO/MI take a bit longer when done during an appointment, so add an extra 10 min for the calculation
                        minforslots = minforslots + 10
                    End If
                End If
            Else
                'Client has slot override, so calculate from there and ignore LY's minutes entirely
                minforslots = minforslots + (.NumApptSlotsToUse * DB_GetSetting(ActiveDBInstance, SETTING_SLOTBREAKPOINT))
            End If

            'Regardless of all the above, we're still calculating total minutes to show on the appointment as a reference
            If .LastYear_MinutesToComplete <> NullLong Then totalminutes = totalminutes + .LastYear_MinutesToComplete
        End With
    Next i

    'Total minutes to show on appointment
    CHOS_NumMinutes = totalminutes

    'Actual slots to use, depending on the calculation above and possibly modified on the fly by the user
    'Examples if SETTING_????? == 40:
    '    39 minutes = 1 slot
    '    40 minutes = 1 slot
    '    41 minutes = 2 slots
    '    79 minutes = 2 slot
    '    80 minutes = 2 slot
    '    81 minutes = 3 slots
    CHOS_NumSlots = Int((minforslots - 1) / DB_GetSetting(ActiveDBInstance, SETTING_SLOTBREAKPOINT)) + 1
    If CHOS_NumSlots = 0 Then CHOS_NumSlots = 1

    CHOS_NumSlotsBeforeOverride = CHOS_NumSlots
    CHOS_NumSlots_Overridden = False

    CHOS_UpdateTotal
    tabSchedule.ChangeScheduleMode sCreate
End If

'Show or hide the listbox
Dim b As Boolean
b = (CHOS_lstClients.ListCount > 0)
CHOS_lstClients.Visible = b
CHOS_lblApptInfo.Visible = b
CHOS_lblClients.Visible = b

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_CalculateTotal", Err
End Sub

'EHT=Standard
Public Sub CHOS_Clear()
On Error GoTo ERR_HANDLER

CHOS_lstClients.Clear
CHOS_CalculateTotal

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_Clear", Err
End Sub

'EHT=Standard
Public Sub CHOS_Remove(ByVal cindex&)
On Error GoTo ERR_HANDLER

'Caution: cindex is index into CHOS_lstClients

Dim i&, cID&
If cindex < 0 Then Exit Sub

CHOS_lstClients.RemoveItem cindex
If CHOS_lstClients.ListCount = 0 Then tabSchedule.ChangeScheduleMode sView

CHOS_CalculateTotal
tabSearch.lstResults.Repaint

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_Remove", Err
End Sub

'EHT=Standard
Public Sub CHOS_UpdateTotal()
On Error GoTo ERR_HANDLER

Dim t$
If CHOS_NumSlots_Overridden Then
    't$ = " *"
    t$ = " (overridden)"
Else
    't$ = ""
    t$ = " (change with +/-)"
End If
CHOS_lblApptInfo.Caption = "Total: " & CHOS_NumMinutes & " " & FormatNumApptSlots(CHOS_NumSlots) & t$

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_UpdateTotal", Err
End Sub

'EHT=Standard
Function CHOS_IsChosen(cID&) As Boolean
On Error GoTo ERR_HANDLER

Dim a&, ub&
ub = CHOS_lstClients.ListCount - 1
For a = 0 To ub
    If cID = CHOS_lstClients.ItemData(a) Then
        CHOS_IsChosen = True
        Exit Function
    End If
Next a

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CHOS_IsChosen", Err
End Function
