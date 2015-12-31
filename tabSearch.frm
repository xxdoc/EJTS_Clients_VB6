VERSION 5.00
Begin VB.Form tabSearch 
   BorderStyle     =   0  'None
   ClientHeight    =   6000
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   12225
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
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   815
   ShowInTaskbar   =   0   'False
   Begin EJTSClients.CustomListbox lstResults 
      Height          =   1000
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1000
      _ExtentX        =   0
      _ExtentY        =   0
      DisplayMode     =   1
      MultiSel        =   0   'False
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   2760
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
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
      Height          =   405
      Left            =   8760
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      ToolTipText     =   "Search syntax help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblCount 
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
      Height          =   405
      Left            =   6840
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu menClient 
      Caption         =   "Client"
      Visible         =   0   'False
      Begin VB.Menu menClient_Title 
         Caption         =   "== Brubaker, Richard A & Bernadette E =="
         Enabled         =   0   'False
      End
      Begin VB.Menu menClient_Sep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menClientEdit 
         Caption         =   "&Edit...               Ctrl+Enter"
      End
      Begin VB.Menu menClientPost 
         Caption         =   "&Post...                      Shift+Enter"
      End
      Begin VB.Menu menClientGotoAppt 
         Caption         =   "&Goto Appointment    Shift+Right"
      End
      Begin VB.Menu menClient_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu menClientMarkDO 
         Caption         =   "&Dropped Off"
      End
      Begin VB.Menu menClientMarkMI 
         Caption         =   "&Mailed In"
      End
      Begin VB.Menu menClientMarkINC 
         Caption         =   "&Incomplete"
      End
      Begin VB.Menu menClientMarkRelBefPmt 
         Caption         =   "&Released Before Payment"
      End
      Begin VB.Menu menClientMarkPaid 
         Caption         =   "P&aid"
      End
      Begin VB.Menu menClientMarkExtension 
         Caption         =   "E&xtension"
      End
      Begin VB.Menu menClient_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu menClientGotoML 
         Caption         =   "Goto Mailing &List"
      End
   End
End
Attribute VB_Name = "tabSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ITab
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Public Enum enumValueType
    tLong = 1
    tString
    tFlags
End Enum
Public Enum enumField
    dID
    dNameALL            'Checks both last names, both first names, and both nicknames
    dNameLast           'Checks both last names
    dNameFirstAndNick   'Checks both first names and both nicknames
    dPhoneALL           'Checks all 3 phone numbers
    dAddressStreet
    dAddressCity
    dAddressState
    dAddressZipCode
    dEmailAddressALL    'Checks both email addresses
    dNotes
    dNumApptSlotsToUse
    dFlags
    dLastYear_MinutesToComplete
    dLastYear_PrepFee
    dLastYear_Flags
    dCompletionDate
    dMinutesToComplete
    dStateList
    dPrepFee
    dMoneyOwed
    dResultAGI
    dResultFederal
    dResultState
    dOpNotes
End Enum
Private Const EnumField_DATAITEMUBOUND = 25 - 1
Public Enum enumOperator
    oEqual 'For flags, means flag set
    oNotEq 'For flags, means flag not set
    oGT
    oLT
    oGTEq
    oLTEq
    oLike  'For string filters
    oNotLike
End Enum
Private Type enumFilter
    Filter_OrOperator As Boolean     'And' is default
    Field As enumField
    Operator As enumOperator
    Value_Long As Long
    Value_String As String
    Value_Flag() As Long
    Value_FlagSet() As Boolean
    Value_FlagCount As Long
    ValueType As enumValueType
End Type
Private Type enumDefinition
    IsSimpleSearch As Boolean
    SimpleSearchStringUCase As String
    NotOperator As Boolean
    Filters() As enumFilter
    FilterCount As Long
End Type
Private Type enumSyntaxItem
    Term As String
    Value As Long
End Type

Public SkipChangeEvents As Boolean
Private CurrentSearch As enumDefinition
Private SyntaxTable_Fields() As enumSyntaxItem
Private SyntaxTable_Flags() As enumSyntaxItem
Private mSearchCount&

Private Sub Form_Load()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

Private Function ITab_CreateGDIObjects() As Boolean
End Function

Private Function ITab_InitializeAfterDBLoad() As Boolean
End Function

Private Sub ITab_AfterTabShown()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "ITab_AfterTabShown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim t$(), n$(), a&, b&, stc&

t$ = Split(DB_GetSetting(ActiveDBInstance, "GLOBAL_SearchSyntax_Fields"), SEP1)
stc = 0
For a = 0 To UBound(t$)
    n$ = Split(t$(a), SEP2)
    For b = 0 To UBound(n$)
        ReDim Preserve SyntaxTable_Fields(stc)
        SyntaxTable_Fields(stc).Term = n$(b)
        SyntaxTable_Fields(stc).Value = a
        stc = stc + 1
    Next b
Next a

t$ = Split(DB_GetSetting(ActiveDBInstance, "GLOBAL_SearchSyntax_Flags"), SEP1)
stc = 0
For a = 0 To UBound(t$)
    n$ = Split(t$(a), SEP2)
    For b = 0 To UBound(n$)
        ReDim Preserve SyntaxTable_Flags(stc)
        SyntaxTable_Flags(stc).Term = n$(b)
        SyntaxTable_Flags(stc).Value = (2 ^ a)
        stc = stc + 1
    Next b
Next a

PopulateCboSpecialSearch

CLEAN_UP:
    If ERR_COUNT > 0 Then
        Erase SyntaxTable_Fields
        Erase SyntaxTable_Flags
    End If

'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub ITab_SetDefaultFocus()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "ITab_SetDefaultFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetFocusWithoutErr lstResults

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

txtSearch.Move 0, 0, Me.ScaleWidth - lblHelp.Width - lblCount.Width - 16
lblCount.Move txtSearch.Left + txtSearch.Width + 8, 0
lblHelp.Move lblCount.Left + lblCount.Width + 8, 0
lstResults.Move 0, txtSearch.Height + 5, Me.ScaleWidth
lstResults.Height = Me.ScaleHeight - lstResults.Top
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabSearch" & "." & "Form_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabSearch" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

Public Sub txtSearch_Change()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "txtSearch_Change": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If SkipChangeEvents Then Exit Sub

Dim a&, b As Boolean
'If text matches any of the special searches, then select it in cboSpecialSearch
For a = 0 To ActiveDBInstance.SpecialSearches_Count - 1
    If LCase$(ActiveDBInstance.SpecialSearches(a).SearchString) = LCase$(txtSearch.Text) Then
        SkipChangeEvents = True
        frmMain.SRCH_cboSpecialSearch.ListIndex = a
        SkipChangeEvents = False
        b = True
    End If
Next a
If Not b Then
    'if not, select nothing
    SkipChangeEvents = True
    frmMain.SRCH_cboSpecialSearch.ListIndex = -1
    SkipChangeEvents = False
End If

'Do search
DoSearch
UpdateTabAsterisk

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lblHelp_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lblHelp_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, b&, t$, fieldnames$(EnumField_DATAITEMUBOUND), flagnames$(ClientFlags_DATAITEMUBOUND)

For a = 0 To UBound(SyntaxTable_Fields)
    b = SyntaxTable_Fields(a).Value
    If fieldnames$(b) <> "" Then fieldnames$(b) = fieldnames$(b) & ", "
    fieldnames$(b) = fieldnames$(b) & SyntaxTable_Fields(a).Term
Next a
For a = 0 To UBound(SyntaxTable_Flags)
    b = Log(SyntaxTable_Flags(a).Value) / Log(2)
    If flagnames$(b) <> "" Then flagnames$(b) = flagnames$(b) & ", "
    flagnames$(b) = flagnames$(b) & SyntaxTable_Flags(a).Term
Next a
     t$ = "Basic syntax:" & vbCrLf & _
          "    filter=value" & vbCrLf & _
          "    filter=""value"" (if contains spaces)" & vbCrLf & _
          "Flag syntax:" & vbCrLf & _
          "    flags=+NN-C-EF" & vbCrLf & _
          "Number & Date operators: = <> < > <= >=" & vbCrLf & _
          "String operators: = <> ~ !~ (~ allows */? wildcards)" & vbCrLf & _
          "Flag operators: =" & vbCrLf & _
          vbCrLf & _
          "############ Field names: #############" & vbCrLf
For a = 0 To UBound(fieldnames$)
    t$ = t$ & fieldnames$(a) & vbCrLf
Next a
t$ = t$ & vbCrLf & _
          "############ Flag names: #############" & vbCrLf
For a = 0 To UBound(flagnames$)
    t$ = t$ & UCase$(flagnames$(a)) & vbCrLf
Next a

ShowInfoMsg t$

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lstResults_GotFocus()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_GotFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SelectFirstItemIfNoSelection lstResults

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

'[Mark] should be Private
Public Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim i&, cID&
frmMain.Form_KeyDown KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

frmMain.IdleSetAction

Select Case KeyCode
Case vbKeyReturn
    KeyCode = 0
    Select Case Shift
    Case 0, Is > 1000 '[Mark]
        If DEBUGMODE Then
            If ActiveDBInstance.IsWriteable Then
                cID = Shift
                'i = tabSearch.lstResults.ListIndex
                'If i < 0 Then Exit Sub
                'cID = tabSearch.lstResults.ItemClientID(i)  'Get client ID
                'If cID = LB_ERR Then Exit Sub      'Separator item, skip
                MsgBox "adding " & cID '[Mark]
                frmMain.CHOS_Add cID, True
                frmMain.CHOS_CalculateTotal
                lstResults.Repaint
                tabSchedule.ChangeScheduleMode sCreate
            End If
        Else
            If ActiveDBInstance.IsWriteable Then
                i = tabSearch.lstResults.ListIndex
                If i < 0 Then Exit Sub
                cID = tabSearch.lstResults.ItemClientID(i)  'Get client ID
                If cID = LB_ERR Then Exit Sub      'Separator item, skip
                frmMain.CHOS_Add cID, True
            End If
        End If
    Case vbShiftMask
        menClientPost_Click
    Case vbCtrlMask
        menClientEdit_Click
    End Select
    
Case vbKeyLeft, vbKeyRight
    If Shift = vbShiftMask Then
        KeyCode = 0
        menClientGotoAppt_Click
    End If

Case vbKeyUp
    If lstResults.ListIndex = 0 Then
        SetFocusWithoutErr txtSearch
        KeyCode = 0
    End If

Case vbKeySpace
    KeyCode = 0
    If lstResults.ListIndex >= 0 Then PopupClientMenu lstResults.ListIndex, False
    
Case vbKeyBack
    KeyCode = 0
    If Len(txtSearch.Text) > 0 Then txtSearch.Text = Left$(txtSearch.Text, Len(txtSearch.Text) - 1)

Case Else
    If IsLetterKey(KeyCode) Then
        PutKeyCodeIntoTextbox txtSearch, KeyCode, False
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

Private Sub lstResults_KeyUp(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

frmMain.Form_KeyUp KeyCode, Shift: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lstResults_KeyPressByCode(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_KeyPressByCode": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

frmMain.Form_KeyPress KeyCode: If KeyCode = 0 Then Exit Sub   'Pass it to the parent form first, Exit if form cancelled the event

frmMain.IdleSetAction

Select Case KeyCode
Case 33, 44, 60, 61, 62, 126
    PutKeyAsciiIntoTextbox txtSearch, KeyCode, False
    SetFocusWithoutErr txtSearch
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

Private Sub lstResults_TabToNextControl(Reverse As Boolean)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_TabToNextControl": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

TabToNextControl Me, False, Reverse

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_MouseDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Select item under mouse
Dim i&
i = SendMessage(lstResults.hListBox, LB_ITEMFROMPOINT, 0, X + (Y * &H10000))
If i > &HFFFF& Then
    lstResults.ListIndex = -1
Else
    i = (i And &HFFFF&)
    If Button = vbRightButton Then
        lstResults.ListIndex = i    'Listbox only does this for left click on a valid item
        
        'Popup menu
        PopupClientMenu lstResults.ListIndex, True
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

Private Sub lstResults_DblClick()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "lstResults_DblClick": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

menClientEdit_Click

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientEdit_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientEdit_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientEdit.Enabled Then Exit Sub

Dim frm As frmClientEditPost, cID&
'Don't check .Enabled, because sometimes this code is called without showing the menu first
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item
Set frm = New frmClientEditPost
If frm.Form_Show(cID, fEdit) Then   'This will mark changed if necessary
    frmMain.DayTotal_Update
    lstResults.Repaint
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

Private Sub menClientPost_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientPost_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientPost.Enabled Then Exit Sub

Dim frm As frmClientEditPost, cID&, cindex&
'Don't check .Enabled, because sometimes this code is called without showing the menu first
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item
cindex = DB_FindClientIndex(ActiveDBInstance, cID)
If cindex < 0 Then Exit Sub
If Flag_IsSet(ActiveDBInstance.Clients(cindex).c.Flags, CompletedReturn) Then Exit Sub
Set frm = New frmClientEditPost
If frm.Form_Show(cID, fPost) Then    'This will mark changed if necessary
    frmMain.DayTotal_Update
    lstResults.Repaint
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

Private Sub menClientGotoAppt_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientGotoAppt_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientGotoAppt.Enabled Then Exit Sub

Dim cID&, d As Date, aindex&
'Don't check .Enabled, because sometimes this code is called without showing the menu first
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item
aindex = DB_GetClientAppt(ActiveDBInstance, cID, d)
If aindex < 0 Then Exit Sub
MouseNullZone_Set 10
tabSchedule.ShowDate d
If DB_DayWithinBitmapRange(ActiveDBInstance, ActiveDBInstance.Appointments(aindex).ApptDate) Then
    tabSchedule.MoveShapeAndSetStyle ActiveDBInstance.Appointments(aindex).ApptDate - tabSchedule.ViewStartDate, ActiveDBInstance.Appointments(aindex).ApptTimeSlot, ActiveDBInstance.Appointments(aindex).NumSlots, Style_ShowAppt
    FlashStopTime = Timer + FlashDuration
    tabSchedule.tmrFlashAppt.Enabled = True
    frmMain.ChangeCurTab vSchedule, False
Else
    ShowErrorMsg "Appointment day not within appointment bitmap range!"
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

Private Sub menClientMarkDO_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientMarkDO_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientMarkDO.Enabled Then Exit Sub

Dim cID&, cindex&
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

cindex = DB_FindClientIndex(ActiveDBInstance, cID)
With ActiveDBInstance.Clients(cindex).c
    If Flag_IsSet(.Flags, DroppedOff) Then
        .Flags = Flag_Remove(.Flags, DroppedOff)
        AddOpNote .OpNotes, "Removed flag: DO"
        tabLogFile.WriteLine "Marked NOT DO: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    Else
        .Flags = Flag_Remove(.Flags, HadAppointment Or MailedIn) Or DroppedOff
        AddOpNote .OpNotes, "Dropped off"
        tabLogFile.WriteLine "Marked DO: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    End If
    ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
End With
lstResults.Repaint
frmMain.SetChangedFlagAndIndication

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientMarkMI_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientMarkMI_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientMarkMI.Enabled Then Exit Sub

Dim cID&, cindex&
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

cindex = DB_FindClientIndex(ActiveDBInstance, cID)
With ActiveDBInstance.Clients(cindex).c
    If Flag_IsSet(.Flags, MailedIn) Then
        .Flags = Flag_Remove(.Flags, MailedIn)
        AddOpNote .OpNotes, "Removed flag: MI"
        tabLogFile.WriteLine "Marked NOT MI: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    Else
        .Flags = Flag_Remove(.Flags, HadAppointment Or DroppedOff) Or MailedIn
        AddOpNote .OpNotes, "Mailed in"
        tabLogFile.WriteLine "Marked MI: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    End If
    ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
End With
lstResults.Repaint
frmMain.SetChangedFlagAndIndication

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientMarkINC_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientMarkINC_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientMarkINC.Enabled Then Exit Sub

Dim cID&, cindex&
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

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
lstResults.Repaint
frmMain.SetChangedFlagAndIndication

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientMarkRelBefPmt_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientMarkRelBefPmt_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientMarkRelBefPmt.Enabled Then Exit Sub

Dim cID&, cindex&
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

cindex = DB_FindClientIndex(ActiveDBInstance, cID)
With ActiveDBInstance.Clients(cindex).c
    If Flag_IsSet(.Flags, ReleasedBeforePayment) Then
        .Flags = Flag_Remove(.Flags, ReleasedBeforePayment)
        AddOpNote .OpNotes, "Removed flag: RelBefPmt"
        tabLogFile.WriteLine "Marked NOT RelBefPmt: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    Else
        .Flags = .Flags Or ReleasedBeforePayment
        AddOpNote .OpNotes, "Released before payment"
        tabLogFile.WriteLine "Marked RelBefPmt: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    End If
    ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
End With
lstResults.Repaint
frmMain.SetChangedFlagAndIndication

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientMarkPaid_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientMarkPaid_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientMarkPaid.Enabled Then Exit Sub

Dim cID&, cindex&
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

cindex = DB_FindClientIndex(ActiveDBInstance, cID)
With ActiveDBInstance.Clients(cindex).c
    .MoneyOwed = NullLong
    AddOpNote .OpNotes, "Paid"
    tabLogFile.WriteLine "Marked Paid: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
End With
lstResults.Repaint
frmMain.SetChangedFlagAndIndication

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientMarkExtension_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientMarkExtension_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientMarkExtension.Enabled Then Exit Sub

Dim cID&, cindex&
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

cindex = DB_FindClientIndex(ActiveDBInstance, cID)
With ActiveDBInstance.Clients(cindex).c
    If Flag_IsSet(.Flags, Extension) Then
        .Flags = Flag_Remove(.Flags, Extension)
        AddOpNote .OpNotes, "Removed flag: EXT"
        tabLogFile.WriteLine "Marked NOT Extension: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    Else
        .Flags = .Flags Or Extension
        AddOpNote .OpNotes, "Extension"
        tabLogFile.WriteLine "Marked Extension: " & FormatClientName(fLog, ActiveDBInstance.Clients(cindex).c)
    End If
    ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
End With
lstResults.Repaint
frmMain.SetChangedFlagAndIndication

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menClientGotoML_Click()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "menClientGotoML_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menClientGotoML.Enabled Then Exit Sub

Dim cID&, a&, b&
'Don't check .Enabled, because sometimes this code is called without showing the menu first
cID = lstResults.SelectedClientID
If cID = LB_ERR Then Exit Sub    'Separator item

frmMain.ChangeCurTab vMailingList, False

For a = 1 To 3
    With tabMailingList.lstSection(a)
        For b = 0 To .ListCount - 1
            If .ItemClientID(b) = cID Then
                SetFocusWithoutErr tabMailingList.lstSection(a)
                .ListIndex = b
                .TopIndex = b
                Exit Sub
            End If
        Next b
    End With
Next a

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "txtSearch_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Select Case KeyCode
Case vbKeyDown
    'Select first client (>=0)
    Dim a&
    For a = 0 To lstResults.ListCount - 1
        If lstResults.ItemClientID(a) >= 0 Then
            lstResults.ListIndex = a
            SetFocusWithoutErr lstResults
            Exit For
        End If
    Next a
    KeyCode = 0
Case vbKeyUp
    KeyCode = 0
Case vbKeyF1
    KeyCode = 0
    lblHelp_Click
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

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "txtSearch_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Select Case KeyAscii
Case vbKeyReturn, vbKeyEscape
    KeyAscii = 0    'Stop the beep
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

Sub ClearAll()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "ClearAll": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SkipChangeEvents = True
frmMain.SRCH_cboSpecialSearch.ListIndex = -1
txtSearch.Text = ""
Erase CurrentSearch.Filters
CurrentSearch.FilterCount = 0
lstResults.Clear
lblCount.Caption = "Count: 0"
UpdateTabAsterisk
SkipChangeEvents = False

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub PopulateCboSpecialSearch()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "PopulateCboSpecialSearch": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&
frmMain.SRCH_cboSpecialSearch.Clear
For a = 0 To ActiveDBInstance.SpecialSearches_Count - 1
    frmMain.SRCH_cboSpecialSearch.AddItem ActiveDBInstance.SpecialSearches(a).DisplayName
Next a

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Sub PopupClientMenu(li&, showmenuatcursor As Boolean)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "PopupClientMenu": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim cID&, cindex&, aindex&, d As Date, comp As Boolean
cID = tabSearch.lstResults.ItemClientID(li)
If cID >= 0 Then    'Valid item (<0 is a separator)
    cindex = DB_FindClientIndex(ActiveDBInstance, cID)
    With ActiveDBInstance.Clients(cindex).c
        comp = Flag_IsSet(.Flags, CompletedReturn) Or Flag_IsSet(.Flags, NoNeedToFile)
        'Show Client menu
        menClient_Title.Caption = "== " & Replace(FormatClientName(fSchedulePct, ActiveDBInstance.Clients(cindex).c), "&", "&&") & " =="
        'If SRCH_CurrentSearchDisplayMode = sPhone Then
        '    'We've already done the searching, in this case
        '    menClientGotoAppt.Enabled = (.Temp_ApptDate <> "")
        'Else
            'Find client's appointment
            aindex = DB_GetClientAppt(ActiveDBInstance, cID, d)
            menClientGotoAppt.Enabled = (aindex >= 0)
        'End If
        menClientPost.Enabled = Not comp
        menClientMarkINC.Checked = Flag_IsSet(.Flags, PartiallyComplete)
        menClientMarkINC.Enabled = (Not comp) And (ActiveDBInstance.IsWriteable)
        menClientMarkDO.Checked = Flag_IsSet(.Flags, DroppedOff)
        menClientMarkDO.Enabled = (Not comp) And (ActiveDBInstance.IsWriteable)
        menClientMarkMI.Checked = Flag_IsSet(.Flags, MailedIn)
        menClientMarkMI.Enabled = (Not comp) And (ActiveDBInstance.IsWriteable)
        menClientMarkRelBefPmt.Checked = Flag_IsSet(.Flags, ReleasedBeforePayment)
        menClientMarkRelBefPmt.Enabled = (.MoneyOwed <> NullLong) And (ActiveDBInstance.IsWriteable)
        menClientMarkPaid.Checked = (.PrepFee > 0) And (.MoneyOwed = NullLong)
        menClientMarkPaid.Enabled = (.MoneyOwed <> NullLong) And (ActiveDBInstance.IsWriteable)
        If (.MoneyOwed <> NullLong) Then
            menClientMarkPaid.Caption = "P&aid (" & FieldToString(.MoneyOwed, mDollar) & " Owed)"
        Else
            menClientMarkPaid.Caption = "P&aid"
        End If
        menClientMarkExtension.Checked = Flag_IsSet(.Flags, Extension)
        menClientMarkExtension.Enabled = (Not comp) And (ActiveDBInstance.IsWriteable)
    End With
    'Must be outside the With block
    If showmenuatcursor Then
        PopupMenu menClient, , , , menClientEdit    'No With blocks!!!
    Else
        PopupMenu menClient, , 500, 250, menClientEdit  'No With blocks!!!
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

Sub RegenerateClientTempData(cindex&)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "RegenerateClientTempData": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim d As Date, aindex&
With ActiveDBInstance.Clients(cindex)
    'Create display string
    .Temp_ParsedName = FindAndMarkSearchTerm(FormatClientName(fSearchResults, .c), CurrentSearch.SimpleSearchStringUCase)
    
    'Create appt list
    aindex = DB_GetClientAppt(ActiveDBInstance, .c.ID, d)
    If aindex >= 0 Then
        .Temp_ApptDate = FormatApptTime2(d)
        .Temp_ApptPast = (ActiveDBInstance.Appointments(aindex).ApptDate < CLng(Date))
        .Temp_DidntHappen = Flag_IsSet(ActiveDBInstance.Appointments(aindex).Flags, DidntHappen)
    Else
        .Temp_ApptDate = ""
        .Temp_ApptPast = False
        .Temp_DidntHappen = False
    End If
    
    .Temp_RegenerateTempData = False
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

Sub UpdateTabAsterisk()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "UpdateTabAsterisk": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Const t1 = "Search  "
Const t2 = "Search *"
Dim t$
t$ = Replace$(frmMain.TabStrip.Tabs(vSearch + 1).Caption, t2, t1)
If txtSearch.Text <> "" Or lstResults.ListCount > 0 Then
    frmMain.TabStrip.Tabs(vSearch + 1).Caption = Replace$(t$, t1, t2)
Else
    frmMain.TabStrip.Tabs(vSearch + 1).Caption = t$
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

Private Function ClientMatchesFilters(cindex&, asearch As enumDefinition) As Boolean
'errheader>
Const PROC_NAME = "tabSearch" & "." & "ClientMatchesFilters": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim findex&, filterresult As Boolean, a&
Dim l1&, l2&
Dim s1$(5), s1_count&, s2$
Dim f1 As Long

With ActiveDBInstance.Clients(cindex).c
    For findex = 0 To asearch.FilterCount - 1
        filterresult = False
        Select Case asearch.Filters(findex).ValueType
        Case tLong
            l2 = asearch.Filters(findex).Value_Long
            Select Case asearch.Filters(findex).Field
            Case dID
                l1 = .ID
            Case dNumApptSlotsToUse
                l1 = .NumApptSlotsToUse
            Case dLastYear_MinutesToComplete
                l1 = .LastYear_MinutesToComplete
            Case dLastYear_PrepFee
                l1 = .LastYear_PrepFee
            Case dCompletionDate
                l1 = .CompletionDate
            Case dMinutesToComplete
                l1 = .MinutesToComplete
            Case dPrepFee
                l1 = .PrepFee
            Case dMoneyOwed
                l1 = .MoneyOwed
            Case dResultAGI
                l1 = .ResultAGI
            Case dResultFederal
                l1 = .ResultFederal
            Case dResultState
                l1 = .ResultState
            Case Else
                filterresult = False
                GoTo AppendResult
            End Select
            Select Case asearch.Filters(findex).Operator
            Case oEqual
                filterresult = (l1 = l2)
            Case oNotEq
                filterresult = (l1 <> l2)
            Case oGT
                filterresult = (l1 > l2)
            Case oLT
                filterresult = (l1 < l2)
            Case oGTEq
                filterresult = (l1 >= l2)
            Case oLTEq
                filterresult = (l1 <= l2)
            End Select
        Case tString
            s2 = LCase$(asearch.Filters(findex).Value_String)
            Select Case asearch.Filters(findex).Field
            Case dNameALL
                s1(0) = LCase$(.Person1.Last)
                s1(1) = LCase$(.Person2.Last)
                s1(2) = LCase$(.Person1.First)
                s1(3) = LCase$(.Person2.First)
                s1(4) = LCase$(.Person1.Nickname)
                s1(5) = LCase$(.Person2.Nickname)
                s1_count = 6
            Case dNameLast
                s1(0) = LCase$(.Person1.Last)
                s1(1) = LCase$(.Person2.Last)
                s1_count = 2
            Case dNameFirstAndNick
                s1(0) = LCase$(.Person1.First)
                s1(1) = LCase$(.Person2.First)
                s1(2) = LCase$(.Person1.Nickname)
                s1(3) = LCase$(.Person2.Nickname)
                s1_count = 4
            Case dPhoneALL
                s1(0) = LCase$(.Person1.Phone)
                s1(1) = LCase$(.Person2.Phone)
                s1(2) = LCase$(.PhoneHome)
                s1_count = 3
            Case dAddressStreet
                s1(0) = LCase$(.AddressStreet)
                s1_count = 1
            Case dAddressCity
                s1(0) = LCase$(.AddressCity)
                s1_count = 1
            Case dAddressState
                s1(0) = LCase$(.AddressState)
                s1_count = 1
            Case dAddressZipCode
                s1(0) = LCase$(.AddressZipCode)
                s1_count = 1
            Case dEmailAddressALL
                s1(0) = LCase$(.Person1.Email)
                s1(1) = LCase$(.Person2.Email)
                s1_count = 2
            Case dNotes
                s1(0) = LCase$(.Notes)
                s1_count = 1
            Case dStateList
                s1(0) = LCase$(.StateList)
                s1_count = 1
            Case dOpNotes
                s1(0) = LCase$(.OpNotes)
                s1_count = 1
            Case Else
                filterresult = False
                GoTo AppendResult
            End Select
            
            filterresult = False
            Select Case asearch.Filters(findex).Operator
            Case oEqual
                For a = 0 To s1_count - 1
                    filterresult = filterresult Or (s1(a) = s2)
                Next a
            Case oNotEq
                For a = 0 To s1_count - 1
                    filterresult = filterresult Or (s1(a) <> s2)
                Next a
            Case oLike
                For a = 0 To s1_count - 1
                    filterresult = filterresult Or (s1(a) Like s2)
                Next a
            Case oNotLike
                For a = 0 To s1_count - 1
                    filterresult = filterresult Or (Not (s1(a) Like s2))
                Next a
            End Select
        Case tFlags
            Select Case asearch.Filters(findex).Field
            Case dFlags
                f1 = .Flags
            Case dLastYear_Flags
                f1 = .LastYear_Flags
            Case Else
                filterresult = False
                GoTo AppendResult
            End Select
            Select Case asearch.Filters(findex).Operator
            Case oEqual
                filterresult = True   'Assume True, then find a flag that doesn't match
                For a = 0 To asearch.Filters(findex).Value_FlagCount - 1
                    If Flag_IsSet(f1, asearch.Filters(findex).Value_Flag(a)) <> _
                       asearch.Filters(findex).Value_FlagSet(a) Then
                        filterresult = False
                        Exit For
                    End If
                Next a
            End Select
        End Select
        
AppendResult:
        If findex = 0 Then
            ClientMatchesFilters = filterresult
        ElseIf asearch.Filters(findex).Filter_OrOperator Then
            'Or
            ClientMatchesFilters = ClientMatchesFilters Or filterresult
        Else
            'And
            ClientMatchesFilters = ClientMatchesFilters And filterresult
        End If
    Next findex
End With

If asearch.NotOperator Then ClientMatchesFilters = Not ClientMatchesFilters

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Function FindAndMarkSearchTerm$(fs$, stu$)
'errheader>
Const PROC_NAME = "tabSearch" & "." & "FindAndMarkSearchTerm": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Finds every occurence of stu$ (UCase) and inserts BoldSep before and after each one
Dim fsu$, a&, la&, stl&
If Len(stu$) = 0 Then
    FindAndMarkSearchTerm$ = fs$
Else
    fsu$ = UCase$(fs$)
    stl = Len(stu$)
    la = -stl + 1
    Do
        a = InStr(la + stl, fsu$, stu$)
        If a = 0 Then Exit Do
        FindAndMarkSearchTerm$ = FindAndMarkSearchTerm$ & Mid$(fs$, la + stl, a - la - stl) & BoldSep & Mid$(fs$, a, stl) & BoldSep
        la = a
    Loop
    FindAndMarkSearchTerm$ = FindAndMarkSearchTerm$ & Mid$(fs$, la + stl)
End If

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Sub DoSearch()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "DoSearch": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim asearch As enumDefinition, t$, estr$, a&
lstResults.SetRedraw False
lstResults.Clear
mSearchCount = 0
t$ = txtSearch.Text
If Len(t$) > 1 Then
    If ParseSearchString(t$, estr$, asearch) Then
        CurrentSearch = asearch
        RunSearch
        
        'Select first client (>=0)
        For a = 0 To lstResults.ListCount - 1
            If lstResults.ItemClientID(a) >= 0 Then
                lstResults.ListIndex = a
                Exit For
            End If
        Next a
    End If
End If
lstResults.SetRedraw True
lblCount.Caption = "Count: " & mSearchCount

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Function IsSimpleSearchString(t$) As Boolean
'errheader>
Const PROC_NAME = "tabSearch" & "." & "IsSimpleSearchString": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, ca&
For a = 1 To Len(t$)
    ca = Asc(Mid$(t$, a, 1))
    If ((ca = 44) Or (ca = 32) Or ((ca >= 65) And (ca <= 90)) Or ((ca >= 97) And (ca <= 122))) Then
        'Allowable character
    Else
        'Stop
        'This cannot be a simple search string
        Exit Function
    End If
Next a
IsSimpleSearchString = True

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Function ParseSearchString(ByVal s$, ByRef estr$, ByRef asearch As enumDefinition) As Boolean
'errheader>
Const PROC_NAME = "tabSearch" & "." & "ParseSearchString": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim parts$(), a&, m&, b&, c$, ca&, Flags$(), flagcount&, withinquotes As Boolean
Dim nextfilterorop As Boolean
Dim f$, o$, v$
Dim tempsearch As enumDefinition

If IsSimpleSearchString(s$) Then
    tempsearch.IsSimpleSearch = True
    tempsearch.SimpleSearchStringUCase = UCase$(s$)

Else
    'Entire search is done in LCase
    c$ = LCase$(s$)
    'Replace all [today] and [today:*] with formatted date
    s$ = ""
    Do
        a = InStr(m + 1, c$, "[today")
        If a = 0 Then Exit Do
        b = InStr(a + 6, c$, "]")
        If b = 0 Then Exit Do
        If Mid$(c$, a + 6, 1) = ":" Then
            o$ = Mid$(c$, a + 7, b - a - 7)
        Else
            o$ = "m/dd/yyyy"
        End If
        s$ = s$ & Mid$(c$, m + 1, a - m - 1) & Format$(Date, o$)
        m = b
    Loop
    s$ = s$ & Mid$(c$, m + 1)
    
    'Custom Split routine (skips the separator if found within "")
    ReDim parts$(0)
    m = 0
    For a = 1 To Len(s$)
        c$ = Mid$(s$, a, 1)
        If c$ = """" Then withinquotes = Not withinquotes
        If withinquotes Then
            parts$(m) = parts$(m) & c$
        Else
            If c$ = " " Then
                If parts$(m) <> "" Then 'This prevents double spaces from creating a blank entry
                    m = m + 1
                    ReDim Preserve parts$(m)
                End If
            Else
                parts$(m) = parts$(m) & c$
            End If
        End If
    Next a
    
    'Parse parts individually
    For a = 0 To UBound(parts$)
        Select Case parts$(a)
        Case "not"
            If a <> 0 Then
                estr$ = "The 'Not' operator only allowed at beginning of search string."
                Exit Function
            End If
            tempsearch.NotOperator = True
        Case "or"
            If tempsearch.FilterCount = 0 Then
                estr$ = "The 'Or' operator only allowed between filters."
                Exit Function
            End If
            nextfilterorop = True
        Case "and"
            If tempsearch.FilterCount = 0 Then
                estr$ = "The 'Or' operator only allowed between filters."
                Exit Function
            End If
            nextfilterorop = False
        Case Else
            m = 0
            f$ = ""
            o$ = ""
            v$ = ""
            'Separate Field, Operator, and Value
            For b = 1 To Len(parts$(a))
                c$ = Mid$(parts$(a), b, 1)
                ca = Asc(c$)
r:
                Select Case m
                Case 0
                    '0-9, a-z, A-Z
                    If ((ca >= 48) And (ca <= 57)) Or ((ca >= 65) And (ca <= 90)) Or ((ca >= 97) And (ca <= 122)) Then
                        f$ = f$ & c$
                    Else
                        m = 1
                        GoTo r
                    End If
                Case 1
                    If InStr("=<>!~", c$) Then
                        o$ = o$ & c$
                    Else
                        m = 2
                        GoTo r
                    End If
                Case 2
                    v$ = v$ & c$
                End Select
            Next b
            If f$ = "" Then
                estr$ = "Data item name missing:" & vbCrLf & parts$(a)
                Exit Function
            End If
            If o$ = "" Then
                estr$ = "Operator missing:" & vbCrLf & parts$(a)
                Exit Function
            End If
            If v$ = "" Then
                estr$ = "Value to search for missing:" & vbCrLf & parts$(a)
                Exit Function
            End If
            
            'Create filter
            ReDim Preserve tempsearch.Filters(tempsearch.FilterCount)
            With tempsearch.Filters(tempsearch.FilterCount)
                .Filter_OrOperator = nextfilterorop
                nextfilterorop = False  'Set back to 'And' for the next filter
                
                'Lookup field name
                .Field = -1
                For b = 0 To UBound(SyntaxTable_Fields)
                    If SyntaxTable_Fields(b).Term = f$ Then
                        .Field = SyntaxTable_Fields(b).Value
                        Exit For
                    End If
                Next b
                If .Field < 0 Then
                    estr$ = "Invalid field name: '" & f$ & "'"
                    Exit Function
                End If
                
                'Select operator
                Select Case Trim$(o$)
                Case "=", "==":         .Operator = oEqual
                Case "!", "!=", "<>":   .Operator = oNotEq
                Case ">", ">>":         .Operator = oGT
                Case "<", "<<":         .Operator = oLT
                Case ">=", "=>":        .Operator = oGTEq
                Case "<=", "=<":        .Operator = oLTEq
                Case "~", "~=", "~~":   .Operator = oLike
                Case "!~":              .Operator = oNotLike
                Case Else
                    estr$ = "Invalid operator: '" & o$ & "'"
                    Exit Function
                End Select
                
                'Determine what type the value is
                v$ = Trim$(v$)
                'Flags look like strings, so we need to catch it first
                If .Field = dFlags Or .Field = dLastYear_Flags Then
                    .ValueType = tFlags
                    
                    'Clear Flag array
                    flagcount = 0
                    Erase Flags$
                    
                    'Separate individual flags (separated by the + or -, no space)
                    For b = 1 To Len(v$)
                        c$ = Mid$(v$, b, 1)
                        If InStr("+-", c$) > 0 Then
                            ReDim Preserve Flags$(flagcount)
                            flagcount = flagcount + 1
                        End If
                        If flagcount > 0 Then
                            Flags$(flagcount - 1) = Flags$(flagcount - 1) & c$
                        End If
                    Next b
                    If flagcount = 0 Then
                        estr$ = "Flag value empty:" & vbCrLf & parts$(a)
                        Exit Function
                    End If
                    
                    'Init flag array
                    ReDim .Value_Flag(flagcount - 1)
                    ReDim .Value_FlagSet(flagcount - 1)
                    .Value_FlagCount = flagcount
                    
                    'Parse each flag individually
                    For ca = 0 To flagcount - 1
                        'Set FlagSet, +:True, -:False
                        .Value_FlagSet(ca) = (Left$(Flags$(ca), 1) = "+")
                        
                        'Lookup flag name
                        c$ = Mid$(Flags$(ca), 2)
                        .Value_Flag(ca) = -1
                        For b = 0 To UBound(SyntaxTable_Flags)
                            If SyntaxTable_Flags(b).Term = c$ Then
                                .Value_Flag(ca) = SyntaxTable_Flags(b).Value
                                Exit For
                            End If
                        Next b
                        If .Value_Flag(ca) < 0 Then
                            estr$ = "Invalid flag name: '" & c$ & "'"
                            Exit Function
                        End If
                    Next ca
                
                ElseIf IsDate(v$) Then
                    .ValueType = tLong
                    .Value_Long = CLng(CDate(v$))
                
                ElseIf IsNumeric(v$) Then
                    .ValueType = tLong
                    .Value_Long = CLng(v$)
                
                Else    'String, with no quotes
                    .ValueType = tString
                    If (Left$(v$, 1) = """") And (Right$(v$, 1) = """") Then
                        .Value_String = Mid$(v$, 2, Len(v$) - 2)
                    Else
                        .Value_String = v$
                    End If
                    If (.Operator = oLike) Or (.Operator = oNotLike) Then
                        If InStr(.Value_String, "*") = 0 And InStr(.Value_String, "?") = 0 Then
                            'If no wildcards were specified, then assume the user meant 'Contains'
                            .Value_String = "*" & .Value_String & "*"
                        End If
                    End If
                End If
            End With
            tempsearch.FilterCount = tempsearch.FilterCount + 1
        End Select
    Next a
End If

asearch = tempsearch
ParseSearchString = True

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Sub RunSearch()
'errheader>
Const PROC_NAME = "tabSearch" & "." & "RunSearch": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim sl&, a&, p&
Dim ul1$, ul2$, uf1$, uf2$

frmMain.ChangeCurTab vSearch, False

If CurrentSearch.IsSimpleSearch Then
    '#################### Simple Search ####################
    Dim sectioncount(4) As Long
    Dim sectiontitles(4) As String
    sectiontitles(0) = "Last name begins with..."
    sectiontitles(1) = "Last name contains..."
    sectiontitles(2) = "First name / nickname contains..."
    sectiontitles(3) = "Last, First contains..."
    sectiontitles(4) = "Notes contains..."
    
    sl = Len(CurrentSearch.SimpleSearchStringUCase)
    
    mSearchCount = 0
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a).c
            ul1$ = UCase$(.Person1.Last)
            ul2$ = UCase$(.Person2.Last)
            uf1$ = UCase$(.Person1.First)
            uf2$ = UCase$(.Person2.First)
            
            'Last begins
            If (Left$(ul1$, sl) = CurrentSearch.SimpleSearchStringUCase) Or _
               (Left$(ul2$, sl) = CurrentSearch.SimpleSearchStringUCase) Then
                p = 0
                sectioncount(p) = sectioncount(p) + 1
            
            'Last contains
            ElseIf InStr(ul1$ & SEP1 & _
                         ul2$, CurrentSearch.SimpleSearchStringUCase) > 0 Then
                p = 1
                sectioncount(p) = sectioncount(p) + 1
            
            'First/nickname contains
            ElseIf InStr(uf1$ & SEP1 & _
                         UCase$(.Person1.Nickname) & SEP1 & _
                         uf2$ & SEP1 & _
                         UCase$(.Person2.Nickname), CurrentSearch.SimpleSearchStringUCase) > 0 Then
                p = 2
                sectioncount(p) = sectioncount(p) + 1
            
            'Last, First cantains
            ElseIf InStr(ul1$ & ", " & uf1$ & SEP1 & _
                         ul1$ & ", " & uf2$ & SEP1 & _
                         ul2$ & ", " & uf1$ & SEP1 & _
                         ul2$ & ", " & uf2$, CurrentSearch.SimpleSearchStringUCase) > 0 Then
                p = 3
                sectioncount(p) = sectioncount(p) + 1
            
            'Notes contains
            ElseIf InStr(UCase$(.Notes), CurrentSearch.SimpleSearchStringUCase) > 0 Then
                p = 4
                sectioncount(p) = sectioncount(p) + 1
            
            Else
                p = -1
            End If
            
            If p >= 0 Then
                lstResults.AddItem (p * 2) + 1, a
                mSearchCount = mSearchCount + 1
            End If
        End With
    Next a
    
    'Add separators
    For a = 0 To UBound(sectioncount)
        If sectioncount(a) > 0 Then
            lstResults.AddItem a * 2, -1, sectiontitles(a) & " (" & CStr(sectioncount(a)) & ")"
        End If
    Next a

Else
    '#################### Criteria Search ####################
    If CurrentSearch.FilterCount = 0 Then Err.Raise 1, , "Invalid search definition"
    
    mSearchCount = 0
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        If ClientMatchesFilters(a, CurrentSearch) Then
            lstResults.AddItem 0, a
            mSearchCount = mSearchCount + 1
        End If
    Next a

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

