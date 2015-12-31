VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   Caption         =   "EJTS Clients"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
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
   ScaleWidth      =   1008
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstSort 
      Height          =   735
      IntegralHeight  =   0   'False
      Left            =   1560
      Sorted          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
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
      Left            =   13680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton btnNewClient 
      Caption         =   "New Client..."
      Height          =   615
      Left            =   13440
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
      Left            =   9720
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
      Width           =   16575
      _ExtentX        =   29236
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
      Left            =   9960
      TabIndex        =   15
      ToolTipText     =   "Click to recalculate"
      Top             =   480
      Width           =   3135
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
      Left            =   11880
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
      Left            =   9720
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
      Left            =   9960
      TabIndex        =   13
      ToolTipText     =   "Click to recalculate"
      Top             =   120
      Width           =   3135
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

Private Sub Form_Load()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.

If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

Function Form_Show() As Boolean
'errheader>
Const PROC_NAME = "frmMain" & "." & "Form_Show": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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
If DB_Load(DataFilesPath & "EJTSClients" & FileToOpen_Year & ".dat", ActiveDBInstance) Then
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
Else
    ERR_COUNT = ERR_COUNT + 1: GoTo CLEAN_UP
End If
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

CLEAN_UP:
    If ERR_COUNT = 0 Then
        Form_Show = True
    Else
        ActiveDBInstance.Loaded = False
        HidePopupInfo
        Unload Me
    End If
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

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "frmMain" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'These get called from the tab-forms. Even if we're not using them yet, leave them here.
End Sub

Sub Form_KeyPress(KeyAscii As Integer)
'These get called from the tab-forms. Even if we're not using them yet, leave them here.
End Sub

Private Sub Form_DblClick()
'errheader>
Const PROC_NAME = "frmMain" & "." & "Form_DblClick": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not ActiveDBInstance.IsWriteable Then
    ShowErrorMsg "Not available in read-only mode!"
    Exit Sub
End If

Dim a&, t$

'If DEBUGMODE Then
'    a = Val(InputBox("Enter number (1296,1304,1127):"))
'    If a > 0 Then tabSearch.lstResults_KeyDown vbKeyReturn, CInt(a)
'    Exit Sub
'End If

Select Case LCase$(InputBox("Enter debug code:"))
Case "copyfees"
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            If Flag_IsSet(.c.Flags, CompletedReturn) Then
                lstSort.AddItem FormatClientName(fPullFiles, .c)
                lstSort.ItemData(lstSort.NewIndex) = a
            End If
        End With
    Next a
    t$ = "ID" & vbTab & "Name" & vbTab & "LYFee" & vbTab & "CYFee" & vbTab & "CYOwed" & vbCrLf
    For a = 0 To lstSort.ListCount - 1
        With ActiveDBInstance.Clients(lstSort.ItemData(a))
            t$ = t$ & .c.ID & vbTab & FormatClientName(fSearchResults, .c) & vbTab & FieldToString(.c.LastYear_PrepFee, mDollarOrNULL) & vbTab & FieldToString(.c.PrepFee, mDollarOrNULL) & vbTab & FieldToString(.c.MoneyOwed, mDollarOrNULL) & vbCrLf
        End With
    Next a
    Clipboard.Clear
    Clipboard.SetText t$
    MsgBox "Data copied to the clipboard.", vbInformation

Case "dod"
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            lstSort.AddItem FormatClientName(fPullFiles, .c)
            lstSort.ItemData(lstSort.NewIndex) = a
        End With
    Next a
    For a = 0 To lstSort.ListCount - 1
        With ActiveDBInstance.Clients(lstSort.ItemData(a))
            If (.c.Person1.dod <> NullLong) Or (.c.Person2.dod <> NullLong) Then
                t$ = t$ & .c.ID & vbTab & .c.Person1.Last & vbTab & .c.Person1.First & vbTab & .c.Person2.First & vbTab & IIf(.c.Person1.dod = NullLong, "", "D") & vbTab & IIf(.c.Person2.dod = NullLong, "", "D") & vbTab & FormatClientName(fSearchResults, .c) & vbCrLf
            End If
        End With
    Next a
    Clipboard.Clear
    Clipboard.SetText t$

Case "t"
    Dim fh As CMNMOD_CFileHandler
    Dim l$()
    Dim cindex&
    Set fh = OpenFile("C:\0Kenneth\Programming\Visual Basic\Programs\EJTSClients\From Dad\newest oldest.txt", mLineByLine_Input)
    Do Until fh.EndOfFile
        t$ = fh.ReadLine
        l$ = Split(t$, vbTab)
        a = Val(l$(0))
        cindex = DB_FindClientIndex(ActiveDBInstance, a)
        If cindex >= 0 Then
            With ActiveDBInstance.Clients(cindex).c
                a = Val(l$(1))
                If a = 9900 Then a = NullLong
                If .OldestYearFiled = 9900 Then .OldestYearFiled = NullLong
                If .OldestYearFiled <> NullLong Then
                    If .OldestYearFiled <> a Then Stop
                End If
                .OldestYearFiled = a
                
                a = Val(l$(2))
                If a = 9900 Then a = NullLong
                If .NewestYearFiled = 9900 Then .NewestYearFiled = NullLong
                If .NewestYearFiled <> NullLong Then
                    If .NewestYearFiled <> a Then Stop
                End If
                .NewestYearFiled = a
            End With
        Else
            Stop
        End If
    Loop
    fh.CloseFile
    Stop
    
Case "copynames"
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            lstSort.AddItem FormatClientName(fPullFiles, .c)
            lstSort.ItemData(lstSort.NewIndex) = a
        End With
    Next a
    For a = 0 To lstSort.ListCount - 1
        With ActiveDBInstance.Clients(lstSort.ItemData(a))
            t$ = t$ & .c.ID & vbTab & FormatClientName(fPullFiles, .c) & vbTab & FieldToString(.c.OldestYearFiled, mYearOrNULL) & vbTab & FieldToString(.c.NewestYearFiled, mYearOrNULL) & vbCrLf
        End With
    Next a
    Clipboard.Clear
    Clipboard.SetText t$

Case "fixnamecase"
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            .c.Person1.Email = LCase(.c.Person1.Email)
            .c.Person2.Email = LCase(.c.Person2.Email)
        End With
    Next a
    SetChangedFlagAndIndication

Case ""
Case Else
    ShowErrorMsg "Unknown debug code!"
End Select

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub Form_Unload(Cancel As Integer)
'errheader>
Const PROC_NAME = "frmMain" & "." & "Form_Unload": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    If ERR_COUNT > 0 Then Cancel = True
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub pctInitialFocus_GotFocus()
'errheader>
Const PROC_NAME = "frmMain" & "." & "pctInitialFocus_GotFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'If the entire program loses focus and then regains it, pctInitialFocus will get
'   the focus, since it's TabIndex=0. If the focus is passed back to the sub-form,
'   nothing happens. But if focus is changed to a fellow-control, and then passed
'   to the sub-form, it works for some reason.
Dim f As Form
Set f = Tabs(CurTab)
SetFocusWithoutErr pctSecondFocus
SetFocusWithoutErr f

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub tmrDate_Timer()
'errheader>
On Error Resume Next        'ALL ERRORS WILL BE IGNORED IN THIS PROCEDURE
'<errheader

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

Private Sub btnNewClient_Click()
'errheader>
Const PROC_NAME = "frmMain" & "." & "btnNewClient_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub btnSave_Click()
'errheader>
Const PROC_NAME = "frmMain" & "." & "btnSave_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub TabStrip_Click()
'errheader>
Const PROC_NAME = "frmMain" & "." & "TabStrip_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim t As MSComctlLib.Tab
Static lt As MSComctlLib.Tab

Set t = TabStrip.SelectedItem
If Not DontCallChangeCurTab Then ChangeCurTab t.Index - 1, True
If Not lt Is Nothing Then lt.HighLighted = False
t.HighLighted = True
Set lt = t

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub tmrAutoSave_Timer()
'errheader>
On Error Resume Next        'ALL ERRORS WILL BE IGNORED IN THIS PROCEDURE
'<errheader

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

Private Sub tmrPopupInfo_Timer()
'errheader>
On Error Resume Next        'ALL ERRORS WILL BE IGNORED IN THIS PROCEDURE
'<errheader

HidePopupInfo
End Sub

Private Sub CHOS_lstClients_DblClick()
'errheader>
Const PROC_NAME = "frmMain" & "." & "CHOS_lstClients_DblClick": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim i&

SetFocusWithoutErr pctInitialFocus

'Remove selected client from chosen list
i = CHOS_lstClients.ListIndex
If i >= 0 Then CHOS_Remove i

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub CHOS_lstClients_GotFocus()
'errheader>
Const PROC_NAME = "frmMain" & "." & "CHOS_lstClients_GotFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetFocusWithoutErr pctInitialFocus

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub SRCH_cboSpecialSearch_Click()
'errheader>
Const PROC_NAME = "frmMain" & "." & "SRCH_cboSpecialSearch_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub SRCH_lblSpecialSearchEdit_Click()
'errheader>
Const PROC_NAME = "frmMain" & "." & "SRCH_lblSpecialSearchEdit_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim frm As frmEditSearches
Set frm = New frmEditSearches
frm.Form_Show

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub DTOT_lblDayTotal_Click(Index As Integer)
'errheader>
Const PROC_NAME = "frmMain" & "." & "DTOT_lblDayTotal_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

DayTotal_Update

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub ChangeCurTab(ct As enumTabName, FromTabStripEvent As Boolean)
'errheader>
Const PROC_NAME = "frmMain" & "." & "ChangeCurTab": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub SetChangedFlagAndIndication()
'errheader>
Const PROC_NAME = "frmMain" & "." & "SetChangedFlagAndIndication": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not ActiveDBInstance.IsWriteable Then
    'This should never actually happen, since there are other protections in the code
    Err.Raise 1, , "Database has been opened in read-only mode, yet SetChangedFlagAndIndication has been called! Your changes will not actually be saved."
End If

ActiveDBInstance.Changed = True
Me.Caption = Me.Tag & " - CHANGED"
btnSave.Caption = btnSave.Tag & " (*)"
tmrAutoSave.Enabled = False
tmrAutoSave.Enabled = True

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub ClearChangedIndication()
'errheader>
Const PROC_NAME = "frmMain" & "." & "ClearChangedIndication": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Me.Caption = Me.Tag
btnSave.Caption = btnSave.Tag

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub ShowPopupInfo(i$, secondstoshow#)
'errheader>
Const PROC_NAME = "frmMain" & "." & "ShowPopupInfo": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub HidePopupInfo()
'errheader>
Const PROC_NAME = "frmMain" & "." & "HidePopupInfo": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not PopupInfoActive Then Exit Sub
PopupInfoActive = False
tmrPopupInfo.Enabled = False
pctPopupInfo.Visible = False

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub IdleSetAction()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
IdleNextTimeout = Now + (1 / 24 / 60 * 5)
End Sub

Sub IdlePauseTimeout()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
IdleNextTimeout = 0
End Sub

Sub DayTotal_Update()
'errheader>
Const PROC_NAME = "modDayTotal" & "." & "DayTotal_Update": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, cd&, tot&, numappt&, t$

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub CHOS_Add(ByVal cID&, GotoScheduleIfAlreadyChosen As Boolean)
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_Add": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub CHOS_Add2(cID&, cindex&)
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_Add2": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

With ActiveDBInstance.Clients(cindex).c
    CHOS_lstClients.AddItem FieldToString(.LastYear_MinutesToComplete, mNumberOrNULL) & vbTab & FormatNumApptSlots(.NumApptSlotsToUse) & vbTab & FormatClientName(fChosenClients, ActiveDBInstance.Clients(cindex).c)
    CHOS_lstClients.ItemData(CHOS_lstClients.NewIndex) = cID
End With

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub CHOS_CalculateTotal()
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_CalculateTotal": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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
                If .LastYear_MinutesToComplete = NullLong Then
                    'If no LY history and no NumSlots override set, then just assume it's a DA, like the new clients
                    minforslots = minforslots + (2 * 40)
                Else
                    If Flag_IsSet(.LastYear_Flags, DroppedOff) Or Flag_IsSet(.LastYear_Flags, MailedIn) Then
                        'DO/MI take a bit longer when done during an appointment, so add an extra 10 min for the calculation
                        minforslots = minforslots + .LastYear_MinutesToComplete + 10
                    Else
                        minforslots = minforslots + .LastYear_MinutesToComplete
                    End If
                End If
            Else
                minforslots = minforslots + (.NumApptSlotsToUse * 40)
            End If
            
            If .LastYear_MinutesToComplete <> NullLong Then totalminutes = totalminutes + .LastYear_MinutesToComplete
        End With
    Next i
    CHOS_NumMinutes = totalminutes
    CHOS_NumSlots = CalcNumApptSlotsFromMinuteSum(minforslots)
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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub CHOS_Clear()
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_Clear": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

CHOS_lstClients.Clear
CHOS_CalculateTotal

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub CHOS_Remove(ByVal cindex&)
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_Remove": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Caution: cindex is index into CHOS_lstClients

Dim i&, cID&
If cindex < 0 Then ERR_COUNT = ERR_COUNT + 1: GoTo CLEAN_UP

CHOS_lstClients.RemoveItem cindex
If CHOS_lstClients.ListCount = 0 Then tabSchedule.ChangeScheduleMode sView

CHOS_CalculateTotal
tabSearch.lstResults.Repaint

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Public Sub CHOS_UpdateTotal()
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_UpdateTotal": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim t$
If CHOS_NumSlots_Overridden Then
    't$ = " *"
    t$ = " (overridden)"
Else
    't$ = ""
    t$ = " (change with +/-)"
End If
CHOS_lblApptInfo.Caption = "Total: " & CHOS_NumMinutes & " " & FormatNumApptSlots(CHOS_NumSlots) & t$

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Function CHOS_IsChosen(cID&) As Boolean
'errheader>
Const PROC_NAME = "modChosenClients" & "." & "CHOS_IsChosen": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, ub&
ub = CHOS_lstClients.ListCount - 1
For a = 0 To ub
    If cID = CHOS_lstClients.ItemData(a) Then
        CHOS_IsChosen = True
        Exit Function
    End If
Next a

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function
