VERSION 5.00
Begin VB.Form frmExtraChargeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Extra Charge"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExtraCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3233
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
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
      Left            =   233
      TabIndex        =   10
      Top             =   2280
      Width           =   2655
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
      Index           =   4
      Left            =   3360
      TabIndex        =   4
      Tag             =   "24"
      Top             =   1800
      Width           =   1215
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
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Tag             =   "23"
      Top             =   1800
      Width           =   1215
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
      Left            =   120
      TabIndex        =   2
      Tag             =   "31"
      Top             =   1800
      Width           =   1215
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
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Tag             =   "50"
      Top             =   1080
      Width           =   4455
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
      Tag             =   "50"
      Top             =   360
      Width           =   4455
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
      Left            =   3705
      TabIndex        =   12
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Owed:"
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
      Left            =   3360
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Date:"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
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
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Client name:"
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
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmExtraChargeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "frmExtraChargeEdit"

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again
Public TabOrderSetting As String            'This is set in Form_Show

Private Enum FieldName
    fClientName
    fDescription
    fCompletionDate
    fPrepFee
    fMoneyOwed
End Enum

Private EditMode As Boolean
Private CurIndex&

'EHT=Custom
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Function Form_Show(eindex&) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

If (eindex < 0) Or (eindex >= ActiveDBInstance.ExtraCharges_Count) Then
    Err.Raise 1, , "Extra Charge #" & eindex & " not found!"
End If

'Set the tab order
TabOrderSetting = "GLOBAL_TabOrder_ExtraChargeEdit"
SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)

With ActiveDBInstance.ExtraCharges(eindex)
    FieldToTextbox txtField(fClientName), .ClientName
    FieldToTextbox txtField(fDescription), .Description
    FieldToTextbox txtField(fPrepFee), .PrepFee
    FieldToTextbox txtField(fMoneyOwed), .MoneyOwed
    FieldToTextbox txtField(fCompletionDate), .CompletionDate
End With

lblChangeTabOrder.Move Me.ScaleWidth - lblChangeTabOrder.Width - 1, Me.ScaleHeight - lblChangeTabOrder.Height - 1

EditMode = True
CurIndex = eindex
btnSave.Enabled = ActiveDBInstance.IsWriteable
frmMain.IdlePauseTimeout
'-----------------------------------
Me.Show 1, frmMain
'-----------------------------------
frmMain.IdleSetAction
Form_Show = True

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Function

'EHT=Cleanup2
Function Form_ShowNew() As Long
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

' Returns     :Index of new Extra Charge, or -1 if cancelled

'Set the tab order
TabOrderSetting = "GLOBAL_TabOrder_ExtraChargeEdit"
SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)

EditMode = False
CurIndex = -1
btnSave.Enabled = ActiveDBInstance.IsWriteable
'-----------------------------------
Me.Show 1, frmMain
'-----------------------------------
Form_ShowNew = CurIndex

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_ShowNew", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
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

Dim e As ExtraCharge
With e
    FieldFromTextbox txtField(fClientName), .ClientName
    If .ClientName = "" Then
        ShowErrorMsg "Missing client name!"
        SetFocusWithoutErr txtField(fClientName)
        Exit Sub
    End If

    FieldFromTextbox txtField(fDescription), .Description

    FieldFromTextbox txtField(fCompletionDate), .CompletionDate
    If .CompletionDate = NullLong Then
        ShowErrorMsg "Missing completion date!"
        SetFocusWithoutErr txtField(fCompletionDate)
        Exit Sub
    End If

    FieldFromTextbox txtField(fPrepFee), .PrepFee
    If .PrepFee = NullLong Then
        ShowErrorMsg "Missing prep fee!"
        SetFocusWithoutErr txtField(fPrepFee)
        Exit Sub
    End If

    FieldFromTextbox txtField(fMoneyOwed), .MoneyOwed
    If .MoneyOwed = 0 Then .MoneyOwed = NullLong

    If EditMode Then
        ActiveDBInstance.ExtraCharges(CurIndex) = e
        tabLogFile.WriteLine "Edited extra charge #" & CurIndex & " (" & .ClientName & ", " & .Description & ")"
    Else
        CurIndex = DB_AddExtraCharge(ActiveDBInstance, e)
        tabLogFile.WriteLine "Created extra charge #" & CurIndex & " (" & .ClientName & ", " & .Description & ")"
    End If
    frmMain.SetChangedFlagAndIndication
End With

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
Case fMoneyOwed
    If Not EditMode Then
        If txtField(fMoneyOwed).Text = "" Then
            txtField(fMoneyOwed).Text = txtField(fPrepFee).Text
        End If
    End If
Case fCompletionDate
    If txtField(fCompletionDate).Text = "" Then
        FieldToTextbox txtField(fCompletionDate), Date
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

