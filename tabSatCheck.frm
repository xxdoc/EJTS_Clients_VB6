VERSION 5.00
Begin VB.Form tabSatCheck 
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
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   3720
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   4
      Left            =   3720
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   7
      Left            =   3720
      TabIndex        =   7
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   5
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
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
      Index           =   6
      Left            =   3720
      TabIndex        =   6
      Top             =   4560
      Width           =   735
   End
   Begin VB.CheckBox chkLastDayOfTaxSeason 
      Caption         =   "Last day of tax season?"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Computer files:"
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
      Index           =   0
      Left            =   840
      TabIndex        =   25
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Filed returns:"
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
      Index           =   1
      Left            =   840
      TabIndex        =   24
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Other state returns:"
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
      Index           =   2
      Left            =   840
      TabIndex        =   23
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Incompletes:"
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
      Index           =   3
      Left            =   840
      TabIndex        =   22
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "DO/MI Incompletes:"
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
      Index           =   4
      Left            =   840
      TabIndex        =   21
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Folders in extension box:"
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
      Index           =   7
      Left            =   840
      TabIndex        =   20
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Bluebooks in office:"
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
      Index           =   5
      Left            =   840
      TabIndex        =   19
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Incompleted SAF on cabinet:"
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
      Index           =   6
      Left            =   840
      TabIndex        =   18
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label PLFL_lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Numbers from Eric:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   17
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label PLFL_lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Numbers from front office:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   16
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 264 (257 CompRet + 7 Inc + 0 IncExt)"
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
      Index           =   0
      Left            =   4560
      TabIndex        =   15
      Top             =   1200
      Width           =   7815
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 249"
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
      Index           =   1
      Left            =   4560
      TabIndex        =   14
      Top             =   1680
      Width           =   7815
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 32"
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
      Index           =   2
      Left            =   4560
      TabIndex        =   13
      Top             =   2160
      Width           =   7815
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 7"
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
      Index           =   3
      Left            =   4560
      TabIndex        =   12
      Top             =   2640
      Width           =   7815
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 249"
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
      Index           =   7
      Left            =   4560
      TabIndex        =   11
      Top             =   5160
      Width           =   7815
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 249"
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
      Index           =   5
      Left            =   4560
      TabIndex        =   10
      Top             =   4200
      Width           =   7815
   End
   Begin VB.Label lblComparison 
      Caption         =   "Should be 249"
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
      Index           =   6
      Left            =   4560
      TabIndex        =   9
      Top             =   4680
      Width           =   7815
   End
End
Attribute VB_Name = "tabSatCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabSatCheck"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Public SkipTxtChange As Boolean

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
SkipTxtChange = True
For a = 0 To txtField.UBound
    txtField(a).Enabled = ActiveDBInstance.IsWriteable
    txtField(a).Text = DB_GetSetting(ActiveDBInstance, "_SatCheck-Txt" & a)
Next a
chkLastDayOfTaxSeason.Value = (Not DB_GetSetting(ActiveDBInstance, "_SatCheck-LastDayOfTaxSeason")) + 1
chkLastDayOfTaxSeason.Enabled = ActiveDBInstance.IsWriteable
SkipTxtChange = False

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

Dim a&, t&, defstate$
Dim db_CompRet&, db_EF&, db_OthSt&, db_Inc&, db_IncExt&, db_SignificantNewClients&, db_Unpaid&, db_RelBefPmt&
Dim eric_TotFiles&, eric_EF&, eric_OthSt&, eric_Inc&, eric_IncDoMi&
Dim office_BB&, office_SAF&, office_ExtFolders&

defstate$ = DB_GetSetting(ActiveDBInstance, "GLOBAL_DefaultState")

For a = 0 To ActiveDBInstance.Clients_Count - 1
    With ActiveDBInstance.Clients(a).c
        If Flag_IsSet(.Flags, CompletedReturn) Then
            db_CompRet = db_CompRet + 1
            If Flag_IsSet(.Flags, EFiled) Then
                db_EF = db_EF + 1
            End If
            If Len(.StateList) > 0 Then
                If .StateList <> defstate$ Then
                    db_OthSt = db_OthSt + 1
                End If
            End If
            If Flag_IsSet(.Flags, NewClient) Then
                If .PrepFee >= 90 Then
                    db_SignificantNewClients = db_SignificantNewClients + 1
                End If
            End If
            If .MoneyOwed <> NullLong Then
                db_Unpaid = db_Unpaid + 1
                If Flag_IsSet(.Flags, ReleasedBeforePayment) Then
                    db_RelBefPmt = db_RelBefPmt + 1
                End If
            End If
        Else
            If Flag_IsSet(.Flags, PartiallyComplete) Then
                db_Inc = db_Inc + 1
            End If
            If Flag_IsSet(.Flags, Extension) Then
                db_IncExt = db_IncExt + 1
            End If
        End If
    End With
Next a

eric_TotFiles = ReadInputTextbox(0)
eric_EF = ReadInputTextbox(1)
eric_OthSt = ReadInputTextbox(2)
eric_Inc = ReadInputTextbox(3)
eric_IncDoMi = ReadInputTextbox(4)

office_BB = ReadInputTextbox(5)
office_SAF = ReadInputTextbox(6)
office_ExtFolders = ReadInputTextbox(7)

If (eric_TotFiles < 0) Or (eric_Inc < 0) Then
    ShowComparison 0, 0, "???"
Else
    If chkLastDayOfTaxSeason.Value = vbChecked Then
        t = db_CompRet + eric_Inc + db_IncExt
    Else
        t = db_CompRet + eric_Inc
    End If
    If eric_TotFiles = t Then
        ShowComparison 0, 1, "Correct"
    Else
        If chkLastDayOfTaxSeason.Value = vbChecked Then
            ShowComparison 0, 2, "Should be " & t & " (" & db_CompRet & " Completed returns + " & eric_Inc & " Incompletes + " & db_IncExt & " Incomplete extensions)"
        Else
            ShowComparison 0, 2, "Should be " & t & " (" & db_CompRet & " Completed returns + " & eric_Inc & " Incompletes)"
        End If
    End If
End If

If eric_EF < 0 Then
    ShowComparison 1, 0, "???"
ElseIf eric_EF = db_EF Then
    ShowComparison 1, 1, "Correct"
Else
    ShowComparison 1, 2, "Should be " & db_EF
End If

If eric_OthSt < 0 Then
    ShowComparison 2, 0, "???"
ElseIf eric_OthSt = db_OthSt Then
    ShowComparison 2, 1, "Correct"
Else
    ShowComparison 2, 2, "Should be " & db_OthSt
End If

If (eric_Inc < 0) Or (eric_IncDoMi < 0) Then
    ShowComparison 3, 0, "???"
Else
    t = db_Inc + eric_IncDoMi
    If eric_Inc = t Then
        ShowComparison 3, 1, "Correct"
    Else
        ShowComparison 3, 2, "Should be " & t
    End If
End If

'Index 4, DO/MI Inc has nothing to check

If office_BB < 0 Then
    ShowComparison 5, 0, "???"
Else
    t = db_Unpaid - db_RelBefPmt
    If office_BB = t Then
        ShowComparison 5, 1, "Correct"
    Else
        ShowComparison 5, 2, "Should be " & t
    End If
End If

If office_SAF < 0 Then
    ShowComparison 6, 0, "???"
ElseIf office_SAF = db_SignificantNewClients Then
    ShowComparison 6, 1, "Correct"
Else
    ShowComparison 6, 2, "Should be " & db_SignificantNewClients
End If

If office_ExtFolders < 0 Then
    ShowComparison 7, 0, "???"
ElseIf office_ExtFolders = db_IncExt Then
    ShowComparison 7, 1, "Correct"
Else
    ShowComparison 7, 2, "Should be " & db_IncExt
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SelectAll tabSatCheck.txtField(0)
SetFocusWithoutErr tabSatCheck.txtField(0)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SetDefaultFocus", Err
End Sub

'EHT=Standard
Private Function ITab_SaveSettingsToDBBeforeClose() As Boolean
On Error GoTo ERR_HANDLER

Dim a&
For a = 0 To txtField.UBound
    DB_SetSetting ActiveDBInstance, "_SatCheck-Txt" & a, txtField(a).Text, sLng
Next a
DB_SetSetting ActiveDBInstance, "_SatCheck-LastDayOfTaxSeason", (chkLastDayOfTaxSeason.Value = 1), sBool

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_SaveSettingsToDBBeforeClose", Err
End Function

'EHT=Standard
Private Function ITab_DestroyGDIObjects() As Boolean
On Error GoTo ERR_HANDLER


Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_DestroyGDIObjects", Err
End Function

'EHT=Standard
Private Sub chkLastDayOfTaxSeason_Click()
On Error GoTo ERR_HANDLER

If SkipTxtChange Then Exit Sub
ITab_AfterTabShown
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkLastDayOfTaxSeason_Click", Err
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
Private Sub txtField_Change(Index As Integer)
On Error GoTo ERR_HANDLER

If SkipTxtChange Then Exit Sub
If Not txtField(Index).Enabled Then Exit Sub
ITab_AfterTabShown
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_Change", Err
End Sub

'EHT=Standard
Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

If KeyCode = vbKeyReturn Then
    TabToNextControl Me, True, (Shift = vbShiftMask)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_KeyDown", Err
End Sub

'EHT=Standard
Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ERR_HANDLER

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0    'Stop the beep
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_KeyPress", Err
End Sub

'EHT=Standard
Function ReadInputTextbox(i&) As Long
On Error GoTo ERR_HANDLER

Dim t$
t$ = txtField(i).Text
If IsNumeric(t$) Then
    ReadInputTextbox = CLng(t$)
Else
    ReadInputTextbox = -1
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ReadInputTextbox", Err
End Function

'EHT=Standard
Sub ShowComparison(i&, t&, c$)
On Error GoTo ERR_HANDLER

With lblComparison(i)
    .Caption = c$
    Select Case t
    Case 0
        .ForeColor = vbWindowText
    Case 1
        .ForeColor = &H8000&    'Green
    Case 2
        .ForeColor = &HC0&      'Red
    End Select
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ShowComparison", Err
End Sub

