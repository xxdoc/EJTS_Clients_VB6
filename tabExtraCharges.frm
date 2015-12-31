VERSION 5.00
Begin VB.Form tabExtraCharges 
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
      Height          =   495
      IntegralHeight  =   0   'False
      ItemData        =   "tabExtraCharges.frx":0000
      Left            =   2040
      List            =   "tabExtraCharges.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstResults 
      Height          =   375
      IntegralHeight  =   0   'False
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   1095
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
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu menItem 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menEdit 
         Caption         =   "&Edit extra charge..."
      End
      Begin VB.Menu menMarkPaid 
         Caption         =   "Mark &paid"
      End
   End
End
Attribute VB_Name = "tabExtraCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ITab
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

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
Const PROC_NAME = "tabExtraCharges" & "." & "ITab_InitializeAfterDBLoad": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetTabStops lstResults.hwnd, 40, 70, 120, 300
btnDelete.Enabled = ActiveDBInstance.IsWriteable
'btnEdit.Enabled = ActiveDBInstance.IsWriteable     'This one is allowed, since frmExtraChargeEdit won't allow saving anyway
btnNew.Enabled = ActiveDBInstance.IsWriteable

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Sub ITab_AfterTabShown()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "ITab_AfterTabShown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Update

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
Const PROC_NAME = "tabExtraCharges" & "." & "ITab_SetDefaultFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

lstResults.Move 0, btnNew.Height + 8, Me.ScaleWidth, Me.ScaleHeight - btnNew.Height - 8
lblCount.Left = Me.ScaleWidth - lblCount.Width
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabExtraCharges" & "." & "Form_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabExtraCharges" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

Private Sub menEdit_Click()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "menEdit_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menEdit.Enabled Then Exit Sub

btnEdit_Click

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub menMarkPaid_Click()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "menMarkPaid_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not menMarkPaid.Enabled Then Exit Sub

Dim i%
i = lstResults.ItemData(lstResults.ListIndex)
With ActiveDBInstance.ExtraCharges(i)
    .MoneyOwed = NullLong
    tabLogFile.WriteLine "Extra Charge entry marked paid: " & .ClientName
End With
ITab_AfterTabShown
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

Private Sub btnNew_Click()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "btnNew_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnNew.Enabled Then Exit Sub

Dim frm As frmExtraChargeEdit, NewIndex&

Set frm = New frmExtraChargeEdit
NewIndex = frm.Form_ShowNew()
If NewIndex >= 0 Then
    'lstResults.AddItem GetText(NewIndex)
    Update
    frmMain.DayTotal_Update
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

Private Sub btnEdit_Click()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "btnEdit_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnEdit.Enabled Then Exit Sub

Dim frm As frmExtraChargeEdit, i&, eindex&

i = lstResults.ListIndex
If i < 0 Then Exit Sub
eindex = lstResults.ItemData(i)
Set frm = New frmExtraChargeEdit
If frm.Form_Show(eindex) Then
    'lstResults.List(i) = GetText$(eindex)
    Update
    frmMain.DayTotal_Update
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

Private Sub btnDelete_Click()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "btnDelete_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnDelete.Enabled Then Exit Sub

Dim i&, eindex&
i = lstResults.ListIndex
If i < 0 Then Exit Sub
eindex = lstResults.ItemData(i)
If MsgBox("Are you sure you want to delete the extra charge for '" & ActiveDBInstance.ExtraCharges(eindex).ClientName & "'?", vbQuestion Or vbYesNo) <> vbYes Then
    Exit Sub
End If
'This must be outside of the with block for it to succeed
DB_RemoveExtraCharge ActiveDBInstance, eindex
Update True
frmMain.DayTotal_Update
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

Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "lstResults_MouseDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Select item under mouse
Dim i&
i = SendMessage(lstResults.hwnd, LB_ITEMFROMPOINT, 0, (X / Screen.TwipsPerPixelX) + ((Y / Screen.TwipsPerPixelY) * &H10000))
If i > &HFFFF& Then
    lstResults.ListIndex = -1
Else
    i = (i And &HFFFF&)
    If Button = vbRightButton Then
        lstResults.ListIndex = i    'Listbox only does this for left click on a valid item
        
        'Popup menu
        i = lstResults.ItemData(i)
        menMarkPaid.Enabled = (ActiveDBInstance.ExtraCharges(i).MoneyOwed <> NullLong) And (ActiveDBInstance.IsWriteable)
        PopupMenu menItem, , , , menEdit
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
Const PROC_NAME = "tabExtraCharges" & "." & "lstResults_DblClick": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

btnEdit_Click

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Function GetText$(eindex&)
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "GetText": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

With ActiveDBInstance.ExtraCharges(eindex)
    If .MoneyOwed <> NullLong Then GetText$ = FieldToString(.MoneyOwed, mDollarOrNULL) & " owed"
    GetText$ = FieldToString(.CompletionDate, mDateAsLongOrNULL) & vbTab & FieldToString(.PrepFee, mDollarOrNULL) & vbTab & GetText$ & vbTab & .ClientName & vbTab & .Description
End With

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Sub Update(Optional DontRestoreSel As Boolean)
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "Update": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, eindex&, lstSelE&, lstTopIndex&
With lstResults
    'Save listbox state
    If .ListIndex >= 0 Then
        lstSelE = .ItemData(.ListIndex)
    Else
        lstSelE = -1
    End If
    lstTopIndex = .TopIndex
    
    'Sort
    lstSort.Clear
    For eindex = 0 To ActiveDBInstance.ExtraCharges_Count - 1
        lstSort.AddItem Format$(ActiveDBInstance.ExtraCharges(eindex).CompletionDate, "yyyy-mm-dd") & ActiveDBInstance.ExtraCharges(eindex).ClientName & ActiveDBInstance.ExtraCharges(eindex).Description
        lstSort.ItemData(lstSort.NewIndex) = eindex
    Next eindex
    
    'Populate listbox
    .Clear
    For a = 0 To lstSort.ListCount - 1
        eindex = lstSort.ItemData(a)
        .AddItem GetText$(eindex)
        .ItemData(.NewIndex) = eindex
        If (eindex = lstSelE) And (Not DontRestoreSel) Then .ListIndex = .NewIndex
    Next a
    UpdateTotal
    
    'Restore listbox state
    .TopIndex = lstTopIndex
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

Sub UpdateTotal()
'errheader>
Const PROC_NAME = "tabExtraCharges" & "." & "UpdateTotal": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

lblCount.Caption = "Count: " & ActiveDBInstance.ExtraCharges_Count

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

