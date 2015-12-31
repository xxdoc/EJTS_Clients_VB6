VERSION 5.00
Begin VB.Form frmBookkeepingEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Bookkeeping"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBookkeepingEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   297
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Tag             =   "31"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Tag             =   "24"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtField 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "23"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label lblChangeTabOrder 
      AutoSize        =   -1  'True
      Caption         =   "Change tab order..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   150
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Completion Date:"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Money Owed:"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Fee:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblDisplayName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "frmBookkeepingEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again
Public TabOrderSetting As String            'This is set in Form_Show

Private Enum FieldName
    fPrepFee
    fCompletionDate
    fMoneyOwed
End Enum

Private IsNewItem As Boolean
Private thisBKIndex&
Private thisMonthIndex&
Private Changed As Boolean

Private Sub Form_Load()
'ANY ERRORS HERE ARE HANDLED BY THE CALLING PROCEDURE
''--..--''--..--''--..--''--..--''--..--''--..--''--.

If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

Function Form_Show(bkindex&, monthindex&) As Boolean
'errheader>
Const PROC_NAME = "frmBookkeepingEdit" & "." & "Form_Show": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Set the tab order
TabOrderSetting = "GLOBAL_TabOrder_BookkeepingEdit"
SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)

thisBKIndex = bkindex
thisMonthIndex = monthindex
With ActiveDBInstance.Bookkeeping(bkindex)
    lblDisplayName.Caption = .DisplayName
    lblMonth.Caption = "(" & MonthName(monthindex + 1) & ")"
End With
With ActiveDBInstance.Bookkeeping(bkindex).Months(monthindex)
    FieldToTextbox txtField(fPrepFee), .PrepFee
    FieldToTextbox txtField(fMoneyOwed), .MoneyOwed
    FieldToTextbox txtField(fCompletionDate), .CompletionDate
    IsNewItem = (.PrepFee = NullLong) And (.CompletionDate = NullLong) And (.MoneyOwed = NullLong)
    If IsNewItem Then Me.Caption = Me.Caption & " - NEW"
End With

SelectAll txtField(fPrepFee)

lblChangeTabOrder.Move Me.ScaleWidth - lblChangeTabOrder.Width - 1, Me.ScaleHeight - lblChangeTabOrder.Height - 1

btnSave.Enabled = ActiveDBInstance.IsWriteable
frmMain.IdlePauseTimeout
'-----------------------------------
Me.Show 1, frmMain
'-----------------------------------
frmMain.IdleSetAction

Form_Show = Changed

CLEAN_UP:
    If ERR_COUNT > 0 Then Unload Me
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "frmBookkeepingEdit" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "frmBookkeepingEdit" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "frmBookkeepingEdit" & "." & "btnSave_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If Not btnSave.Enabled Then Exit Sub

Dim tempbk As BookkeepingMonth

With tempbk
    FieldFromTextbox txtField(fPrepFee), .PrepFee
    FieldFromTextbox txtField(fMoneyOwed), .MoneyOwed
    FieldFromTextbox txtField(fCompletionDate), .CompletionDate
    If .CompletionDate = NullLong Then
        If .PrepFee = NullLong Then
            .MoneyOwed = NullLong
        Else
            ShowErrorMsg "Completion date is not specified"
            SetFocusWithoutErr txtField(fCompletionDate)
            Exit Sub
        End If
    End If
    
    'Write temp copy back to database
    ActiveDBInstance.Bookkeeping(thisBKIndex).Months(thisMonthIndex) = tempbk

    frmMain.SetChangedFlagAndIndication
    tabLogFile.WriteLine "Edited bookkeeping item #" & thisBKIndex & "/" & thisMonthIndex
End With

Changed = True
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
Const PROC_NAME = "frmBookkeepingEdit" & "." & "btnCancel_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

Private Sub txtField_GotFocus(Index As Integer)
'errheader>
Const PROC_NAME = "frmBookkeepingEdit" & "." & "txtField_GotFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Select Case Index
Case fCompletionDate
    If IsNewItem Then
        If txtField(fCompletionDate).Text = "" Then
            If txtField(fPrepFee).Text <> "" Then
                FieldToTextbox txtField(fCompletionDate), Date
                SelectAll txtField(fCompletionDate)
            End If
        End If
    End If
Case fMoneyOwed
    If IsNewItem Then
        If txtField(fMoneyOwed).Text = "" Then
            If txtField(fPrepFee).Text <> "" Then
                txtField(fMoneyOwed).Text = txtField(fPrepFee).Text
                SelectAll txtField(fMoneyOwed)
            End If
        End If
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

Private Sub txtField_LostFocus(Index As Integer)
'errheader>
Const PROC_NAME = "frmBookkeepingEdit" & "." & "txtField_LostFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

LostFocusFormat txtField(Index)

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

Private Sub lblChangeTabOrder_Click()
'errheader>
Const PROC_NAME = "frmBookkeepingEdit" & "." & "lblChangeTabOrder_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim f As frmChangeTabOrder
Set f = New frmChangeTabOrder
f.Form_Show Me

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

