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
Private Const MOD_NAME = "tabExtraCharges"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

'EHT=Custom
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

SetTabStops lstResults.hwnd, 40, 70, 120, 300
btnDelete.Enabled = ActiveDBInstance.IsWriteable
'btnEdit.Enabled = ActiveDBInstance.IsWriteable     'This one is allowed, since frmExtraChargeEdit won't allow saving anyway
btnNew.Enabled = ActiveDBInstance.IsWriteable

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

Update

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr lstResults

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

lstResults.Move 0, btnNew.Height + 8, Me.ScaleWidth, Me.ScaleHeight - btnNew.Height - 8
lblCount.Left = Me.ScaleWidth - lblCount.Width
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
Private Sub menEdit_Click()
On Error GoTo ERR_HANDLER

If Not menEdit.Enabled Then Exit Sub

btnEdit_Click

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menEdit_Click", Err
End Sub

'EHT=Standard
Private Sub menMarkPaid_Click()
On Error GoTo ERR_HANDLER

If Not menMarkPaid.Enabled Then Exit Sub

Dim i%
i = lstResults.ItemData(lstResults.ListIndex)
With ActiveDBInstance.ExtraCharges(i)
    .MoneyOwed = NullLong
    tabLogFile.WriteLine "Extra Charge entry marked paid: " & .ClientName
End With
ITab_AfterTabShown
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menMarkPaid_Click", Err
End Sub

'EHT=Standard
Private Sub btnNew_Click()
On Error GoTo ERR_HANDLER

If Not btnNew.Enabled Then Exit Sub

Dim frm As frmExtraChargeEdit, NewIndex&

Set frm = New frmExtraChargeEdit
NewIndex = frm.Form_ShowNew()
If NewIndex >= 0 Then
    'lstResults.AddItem GetText(NewIndex)
    Update
    frmMain.DayTotal_Update
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnNew_Click", Err
End Sub

'EHT=Standard
Private Sub btnEdit_Click()
On Error GoTo ERR_HANDLER

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

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnEdit_Click", Err
End Sub

'EHT=Standard
Private Sub btnDelete_Click()
On Error GoTo ERR_HANDLER

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

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnDelete_Click", Err
End Sub

'EHT=Standard
Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER

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

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstResults_MouseDown", Err
End Sub

'EHT=Standard
Private Sub lstResults_DblClick()
On Error GoTo ERR_HANDLER

btnEdit_Click

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstResults_DblClick", Err
End Sub

'EHT=Standard
Function GetText$(eindex&)
On Error GoTo ERR_HANDLER

With ActiveDBInstance.ExtraCharges(eindex)
    If .MoneyOwed <> NullLong Then GetText$ = FieldToString(.MoneyOwed, mDollarOrNULL) & " owed"
    GetText$ = FieldToString(.CompletionDate, mDateAsLongOrNULL) & vbTab & FieldToString(.PrepFee, mDollarOrNULL) & vbTab & GetText$ & vbTab & .ClientName & vbTab & .Description
End With

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "GetText", Err
End Function

'EHT=Standard
Sub Update(Optional DontRestoreSel As Boolean)
On Error GoTo ERR_HANDLER

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

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Update", Err
End Sub

'EHT=Standard
Sub UpdateTotal()
On Error GoTo ERR_HANDLER

lblCount.Caption = "Count: " & ActiveDBInstance.ExtraCharges_Count

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UpdateTotal", Err
End Sub

