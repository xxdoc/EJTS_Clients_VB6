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
Private Const MOD_NAME = "frmBookkeepingEdit"

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

'EHT=Custom
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Function Form_Show(bkindex&, monthindex&) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

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

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Function

'EHT=Standard
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

Select Case KeyCode
Case vbKeyReturn
    If Shift = vbCtrlMask Then
        SetFocusWithoutErr btnSave
        btnSave_Click
    Else
        TabToNextControl Me, True, (Shift = vbShiftMask)
    End If
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=Standard
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_HANDLER

Select Case KeyAscii
Case vbKeyReturn
    KeyAscii = 0    'Stop the beep
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyPress", Err
End Sub

'EHT=Standard
Private Sub btnSave_Click()
On Error GoTo ERR_HANDLER

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

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSave_Click", Err
End Sub

'EHT=Standard
Private Sub btnCancel_Click()
On Error GoTo ERR_HANDLER

If Not btnCancel.Enabled Then Exit Sub

Unload Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnCancel_Click", Err
End Sub

'EHT=Standard
Private Sub txtField_GotFocus(Index As Integer)
On Error GoTo ERR_HANDLER

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

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_GotFocus", Err
End Sub

'EHT=Standard
Private Sub txtField_LostFocus(Index As Integer)
On Error GoTo ERR_HANDLER

LostFocusFormat txtField(Index)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_LostFocus", Err
End Sub

'EHT=Standard
Private Sub lblChangeTabOrder_Click()
On Error GoTo ERR_HANDLER

Dim f As frmChangeTabOrder
Set f = New frmChangeTabOrder
f.Form_Show Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lblChangeTabOrder_Click", Err
End Sub

