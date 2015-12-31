VERSION 5.00
Begin VB.Form frmEditSearches 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Custom Search List"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditSearches.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCopyFromFrmMain 
      Caption         =   "&Copy from search box"
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   5
      Top             =   5280
      Width           =   2655
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4080
      Width           =   12135
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   12135
   End
   Begin VB.ListBox lstSearches 
      Height          =   3375
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   12135
   End
End
Attribute VB_Name = "frmEditSearches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private SkipChangeEvents As Boolean

Private Sub Form_Load()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.

If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

Sub Form_Show()
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "Form_Show": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, t$

SetTabStops lstSearches.hwnd, 100
t$ = frmMain.SRCH_cboSpecialSearch.Tag
For a = 0 To ActiveDBInstance.SpecialSearches_Count - 1
    lstSearches.AddItem ActiveDBInstance.SpecialSearches(a).DisplayName & vbTab & ActiveDBInstance.SpecialSearches(a).SearchString
    If ActiveDBInstance.SpecialSearches(a).DisplayName = t$ Then lstSearches.ListIndex = lstSearches.NewIndex
Next a
If ActiveDBInstance.SpecialSearches_Count > 0 Then
    If lstSearches.ListIndex < 0 Then lstSearches.ListIndex = 0
Else
    lstSearches_Click
End If

frmMain.IdlePauseTimeout
'-----------------------------------
Me.Show 1, frmMain
'-----------------------------------
frmMain.IdleSetAction

CLEAN_UP:
    If ERR_COUNT > 0 Then Unload Me
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Select Case KeyCode
Case vbKeyReturn
    If Shift = vbCtrlMask Then
        SetFocusWithoutErr btnSave
        btnSave_Click
    Else
        TabToNextControl Me, True, (Shift = vbShiftMask)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Select Case KeyAscii
Case vbKeyReturn
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

Private Sub btnSave_Click()
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "btnSave_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnSave.Enabled Then Exit Sub

Dim a&, t$()
ActiveDBInstance.SpecialSearches_Count = lstSearches.ListCount
ReDim ActiveDBInstance.SpecialSearches(ActiveDBInstance.SpecialSearches_Count - 1)
For a = 0 To ActiveDBInstance.SpecialSearches_Count - 1
    t$ = Split(lstSearches.List(a), vbTab)
    ActiveDBInstance.SpecialSearches(a).DisplayName = t$(0)
    ActiveDBInstance.SpecialSearches(a).SearchString = t$(1)
Next a

tabSearch.PopulateCboSpecialSearch
frmMain.SetChangedFlagAndIndication
tabSearch.txtSearch_Change

Unload Me

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub btnCancel_Click()
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "btnCancel_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnCancel.Enabled Then Exit Sub

Unload Me

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub btnCopyFromfrmMain_Click()
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "btnCopyFromfrmMain_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnCopyFromFrmMain.Enabled Then Exit Sub

txtSearch(1).Text = tabSearch.txtSearch.Text

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
Const PROC_NAME = "frmEditSearches" & "." & "btnDelete_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnDelete.Enabled Then Exit Sub

Dim a&
a = lstSearches.ListIndex
lstSearches.RemoveItem a
If lstSearches.ListCount > 0 Then
    If a >= lstSearches.ListCount Then a = lstSearches.ListCount - 1
    lstSearches.ListIndex = a
End If
lstSearches_Click
SetFocusWithoutErr lstSearches

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
Const PROC_NAME = "frmEditSearches" & "." & "btnNew_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnNew.Enabled Then Exit Sub

lstSearches.AddItem vbTab
lstSearches.ListIndex = lstSearches.NewIndex
SetFocusWithoutErr txtSearch(0)

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lstSearches_Click()
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "lstSearches_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim t$(), li&, b As Boolean

li = lstSearches.ListIndex
b = li >= 0
btnDelete.Enabled = b
btnCopyFromFrmMain.Enabled = b And (tabSearch.txtSearch.Text <> "")
txtSearch(0).Enabled = b
txtSearch(1).Enabled = b
If b Then
    t$ = Split(lstSearches.List(lstSearches.ListIndex), vbTab)
    txtSearch(0).Text = t$(0)
    txtSearch(1).Text = t$(1)
Else
    txtSearch(0).Text = ""
    txtSearch(1).Text = ""
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

Private Sub lstSearches_DblClick()
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "lstSearches_DblClick": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetFocusWithoutErr txtSearch(1)

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub txtSearch_Change(Index As Integer)
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "txtSearch_Change": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If SkipChangeEvents Then Exit Sub
If lstSearches.ListIndex < 0 Then Exit Sub
lstSearches.List(lstSearches.ListIndex) = Replace(txtSearch(0).Text, vbTab, " ") & vbTab & Replace(Replace(Replace(Replace(txtSearch(1).Text, vbTab, " "), vbCrLf, " "), vbCr, " "), vbLf, " ")

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub txtSearch_GotFocus(Index As Integer)
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "txtSearch_GotFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Index = 0 Then txtSearch(Index).Tag = txtSearch(Index).Text

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub txtSearch_LostFocus(Index As Integer)
'errheader>
Const PROC_NAME = "frmEditSearches" & "." & "txtSearch_LostFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim t$
If Index = 0 Then
    If txtSearch(0).Text <> txtSearch(0).Tag Then
        t$ = lstSearches.List(lstSearches.ListIndex)
        lstSearches.RemoveItem lstSearches.ListIndex
        lstSearches.AddItem t$
        lstSearches.ListIndex = lstSearches.NewIndex
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

