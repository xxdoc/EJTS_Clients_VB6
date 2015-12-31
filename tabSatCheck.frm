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
Implements ITab
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Public SkipTxtChange As Boolean

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
Const PROC_NAME = "tabSatCheck" & "." & "ITab_InitializeAfterDBLoad": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&
SkipTxtChange = True
For a = 0 To txtField.UBound
    txtField(a).Enabled = ActiveDBInstance.IsWriteable
    txtField(a).Text = DB_GetSetting(ActiveDBInstance, "_SatCheck-Txt" & a)
Next a
chkLastDayOfTaxSeason.Value = (Not DB_GetSetting(ActiveDBInstance, "_SatCheck-LastDayOfTaxSeason")) + 1
chkLastDayOfTaxSeason.Enabled = ActiveDBInstance.IsWriteable
SkipTxtChange = False

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
Const PROC_NAME = "tabSatCheck" & "." & "ITab_AfterTabShown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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
Const PROC_NAME = "tabSatCheck" & "." & "ITab_SetDefaultFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SelectAll tabSatCheck.txtField(0)
SetFocusWithoutErr tabSatCheck.txtField(0)

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
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "ITab_SaveSettingsToDBBeforeClose": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&
For a = 0 To txtField.UBound
    DB_SetSetting ActiveDBInstance, "_SatCheck-Txt" & a, txtField(a).Text, sLng
Next a
DB_SetSetting ActiveDBInstance, "_SatCheck-LastDayOfTaxSeason", (chkLastDayOfTaxSeason.Value = 1), sBool

CLEAN_UP:
    'Your code here
'errfooter>
Exit Function
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Function

Private Function ITab_DestroyGDIObjects() As Boolean
End Function

Private Sub chkLastDayOfTaxSeason_Click()
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "chkLastDayOfTaxSeason_Click": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If SkipTxtChange Then Exit Sub
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabSatCheck" & "." & "Form_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabSatCheck" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

Private Sub txtField_Change(Index As Integer)
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "txtField_Change": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If SkipTxtChange Then Exit Sub
If Not txtField(Index).Enabled Then Exit Sub
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

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "txtField_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If KeyCode = vbKeyReturn Then
    TabToNextControl Me, True, (Shift = vbShiftMask)
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

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "txtField_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0    'Stop the beep
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

Function ReadInputTextbox(i&) As Long
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "ReadInputTextbox": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim t$
t$ = txtField(i).Text
If IsNumeric(t$) Then
    ReadInputTextbox = CLng(t$)
Else
    ReadInputTextbox = -1
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

Sub ShowComparison(i&, t&, c$)
'errheader>
Const PROC_NAME = "tabSatCheck" & "." & "ShowComparison": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

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

CLEAN_UP:
    'Your code here
'errfooter>
Exit Sub
ERR_HANDLER:
    If ERR_COUNT >= MAXERRS Then: Err.Raise Err.Number, , Err.Description
    ERR_COUNT = ERR_COUNT + 1: UNHANDLEDERROR PROC_NAME: Resume CLEAN_UP
'<errfooter
End Sub

