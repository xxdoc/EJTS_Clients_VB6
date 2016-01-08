VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EJTS Clients vXXX.XXX.XXX - Choose data file..."
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSetDataFolder 
      Caption         =   "Set Default Folder"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtDataFolder 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Open &read-only"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton btnNewFile 
      Caption         =   "&New File..."
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "&Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ListBox lstDataFiles 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   8175
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "frmStart"

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

'Any calls to DB_GetSetting, DB_SetSetting, or DB_SetDefaultSettingValue should have DontCallSetChangedFlag=True

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Sub Form_Show()
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

SetTabStops lstDataFiles.hwnd, 20, 115

txtDataFolder.Text = DataFilesPath

'Show start form
If DEBUGMODE Then
    Me.Caption = "EJTS Clients vXXX.XXX.XXX - Choose data file..."
    Me.Icon = LoadPicture(AppPath & "DebugMode.ico")
Else
    Me.Caption = "EJTS Clients v" & App.Major & "." & App.Minor & "." & App.Revision & " - Choose data file..."
End If
btnNewFile.Enabled = Not FileToOpen_OpenReadOnly
chkReadOnly.Value = ((Not FileToOpen_OpenReadOnly) + 1)
PopulateListbox
If FileToOpen_Year > 0 Then
    If OpenMainForm Then    'This will call Unload Me if successful
        Exit Sub
    End If
End If
TopMost Me.hwnd
Me.Show

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Sub

'EHT=Standard
Private Sub btnOpen_Click()
On Error GoTo ERR_HANDLER

If Not btnOpen.Enabled Then Exit Sub

Dim f$, li&
li = lstDataFiles.ListIndex
If li >= 0 Then
    f$ = lstDataFiles.List(li)
    FileToOpen_Year = lstDataFiles.ItemData(li)
    OpenMainForm    'This will call Unload Me if successful
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnOpen_Click", Err
End Sub

'EHT=Standard
Private Sub btnNewFile_Click()
On Error GoTo ERR_HANDLER

If Not btnNewFile.Enabled Then Exit Sub

'##############################################################
' This is the 'proforma' function that creates a new database
' using data from the previous year's database, and modifying
' a few things here and there
'##############################################################

Dim a&, b&, s$, OldTaxYear&, NewTaxYear&, TempDBInstance As EJTSClientsDB
Dim OldFile$, NewFile$
Dim Subtitles_Month() As Integer, Subtitles_Day() As Integer, Subtitles_Text() As String, Subtitles_Count As Long

If lstDataFiles.ListCount = 0 Then
    ShowErrorMsg "No files to work with!"
    Exit Sub
End If
OldTaxYear = lstDataFiles.ItemData((lstDataFiles.ListCount - 1))
NewTaxYear = OldTaxYear + 1

'Load last year's database into TempDBInstance
OldFile$ = DataFilesPath & "EJTSClients" & OldTaxYear & ".dat"
NewFile$ = DataFilesPath & "EJTSClients" & NewTaxYear & ".dat"
If FileExists(NewFile$) Then
    ShowErrorMsg "New file '" & NewFile$ & "' already exists!"
    Exit Sub
End If
If Not DB_Load(OldFile$, TempDBInstance) Then Exit Sub
TempDBInstance.FullPath_DB = NewFile$
TempDBInstance.FullPath_Log = DB_GenerateLogfileName(TempDBInstance.FullPath_DB)
TempDBInstance.IsWriteable = True

'Get all the client datasets ready for the new year
For a = 0 To TempDBInstance.Clients_Count - 1
    With TempDBInstance.Clients(a).c
        .LastYear_MinutesToComplete = .MinutesToComplete    'Must be before .MinutesToComplete
        .LastYear_PrepFee = .PrepFee                        'Must be before .PrepFee
        .MinutesToComplete = NullLong
        .PrepFee = NullLong
        .MoneyOwed = NullLong
        .ResultAGI = NullLong
        .ResultFederal = NullLong
        .ResultState = NullLong
        .CompletionDate = NullLong
        .StateList = ""
        .OpNotes = ""
        If Flag_IsSet(.Flags, CompletedReturn) Then         'Must be before .Flags
            .NewestYearFiled = OldTaxYear
            If Flag_IsSet(.Flags, NewClient) Then
                If .OldestYearFiled = NullLong Then .OldestYearFiled = OldTaxYear
            End If
        End If
        .LastYear_Flags = .Flags                            'Must be before .Flags
        .Flags = 0
        If Flag_IsSet(.LastYear_Flags, IncPtnrTrustEstate) Then .Flags = .Flags Or IncPtnrTrustEstate
    End With
Next a

'Clear out the data of all bookkeeping entries (don't delete any entries)
For a = 0 To TempDBInstance.Bookkeeping_Count - 1
    With TempDBInstance.Bookkeeping(a)
        For b = 0 To 11
            .Months(b).CompletionDate = NullLong
            .Months(b).MoneyOwed = NullLong
            .Months(b).PrepFee = NullLong
        Next b
    End With
Next a

'Delete all extra charges
TempDBInstance.ExtraCharges_Count = 0
Erase TempDBInstance.ExtraCharges

'Delete all appointments
TempDBInstance.Appointments_Count = 0
Erase TempDBInstance.Appointments

'Clear out all Saturday Check stuff
For a = 0 To 8
    DB_SetSetting TempDBInstance, "_SatCheck-Txt" & a, 0, sLng, True
Next a
DB_SetSetting TempDBInstance, "_SatCheck-LastDayOfTaxSeason", False, sBool, True

'Set schedule template range A (early January) to equal last year's range C (end of December)
For a = 0 To 6
    s$ = DB_GetSetting(TempDBInstance, "Schedule Template C" & (a + 1) & " (" & WeekdayName(a + 1, False, vbMonday) & ")", , True)
    DB_SetSetting TempDBInstance, "Schedule Template A" & (a + 1) & " (" & WeekdayName(a + 1, False, vbMonday) & ")", s$, sStr, True
Next a

'Save the list of subtitles in a new array by month/day
For a = 0 To TempDBInstance.ApptBitmap_Count - 1
    If Len(TempDBInstance.Subtitles(a)) > 0 Then
        ReDim Preserve Subtitles_Month(Subtitles_Count)
        ReDim Preserve Subtitles_Day(Subtitles_Count)
        ReDim Preserve Subtitles_Text(Subtitles_Count)
        Subtitles_Month(Subtitles_Count) = Month(TempDBInstance.ApptBitmap_StartDate + a)
        Subtitles_Day(Subtitles_Count) = Day(TempDBInstance.ApptBitmap_StartDate + a)
        Subtitles_Text(Subtitles_Count) = TempDBInstance.Subtitles(a)
        Subtitles_Count = Subtitles_Count + 1
    End If
Next a

'Update the bitmap to reflect Jan 1st and Dec 31st of new year (tax year +1)
TempDBInstance.ApptBitmap_StartDate = DateSerial(NewTaxYear + 1, 1, 1)
TempDBInstance.ApptBitmap_Count = DateSerial(NewTaxYear + 1, 12, 31) - TempDBInstance.ApptBitmap_StartDate + 1
ReDim TempDBInstance.ApptBitmap(TempDBInstance.ApptBitmap_Count - 1, Appointment_NumSlotsUB)
ReDim TempDBInstance.Subtitles(TempDBInstance.ApptBitmap_Count - 1)
DB_ClearAndRebuildApptBitmap TempDBInstance

'Now that the bitmap has been updated, restore the subtitles by month/day (no other logic used)
For a = 0 To Subtitles_Count - 1
    'FYI: DateSerial(2015,2,29) will return 3/1/2015, so no error to protect against
    '     However, if last year's 3/1 had a subtitle, it will override the 2/29 (3/1) subtitle
    b = DateSerial(NewTaxYear + 1, Subtitles_Month(a), Subtitles_Day(a)) - TempDBInstance.ApptBitmap_StartDate
    If b >= 0 And b < TempDBInstance.ApptBitmap_Count Then  'This line is just for safety
        TempDBInstance.Subtitles(b) = Subtitles_Text(a)
    End If
Next a

'Save as new filename
If Not DB_Save(TempDBInstance) Then
    ShowErrorMsg "Database save failed!"
    Exit Sub
End If

'Refresh file list
PopulateListbox
SetFocusWithoutErr lstDataFiles

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnNewFile_Click", Err
End Sub

'EHT=Standard
Private Sub btnSetDataFolder_Click()
On Error GoTo ERR_HANDLER

DB_SetSetting ActiveDBInstance, "GLOBAL_DataFolder", DataFilesPath, , True
SetFocusWithoutErr lstDataFiles

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSetDataFolder_Click", Err
End Sub

'EHT=Standard
Private Sub lstDataFiles_DblClick()
On Error GoTo ERR_HANDLER

btnOpen_Click

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstDataFiles_DblClick", Err
End Sub

'EHT=Standard
Private Sub chkReadOnly_Click()
On Error GoTo ERR_HANDLER

FileToOpen_OpenReadOnly = (chkReadOnly.Value = 1)
btnNewFile.Enabled = Not FileToOpen_OpenReadOnly
SetFocusWithoutErr lstDataFiles

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkReadOnly_Click", Err
End Sub

'EHT=Standard
Private Sub txtDataFolder_Change()
On Error GoTo ERR_HANDLER

DataFilesPath = AddTrailingSlash(txtDataFolder.Text)
PopulateListbox

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtDataFolder_Change", Err
End Sub

'EHT=Standard
Private Sub txtDataFolder_LostFocus()
On Error GoTo ERR_HANDLER

If txtDataFolder.Text = "" Then
    txtDataFolder.Text = AppPath & "Data Files\"
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtDataFolder_LostFocus", Err
End Sub

'EHT=Standard
Function OpenMainForm() As Boolean
On Error GoTo ERR_HANDLER

'Load and subclass frmMain (it will hook the listbox parent)
If FileToOpen_Year <> 0 Then
    Set frmMain = New frmMain
    NotTopMost Me.hwnd
    If frmMain.Form_Show Then
        Unload Me
        OpenMainForm = True
    Else
        TopMost Me.hwnd
        Unload frmMain
        FileToOpen_Year = 0
    End If
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "OpenMainForm", Err
End Function

'EHT=Custom
Sub PopulateListbox()
On Error GoTo e
Dim f$, Y&, skipsel As Boolean
lstDataFiles.Enabled = True
lstDataFiles.Clear
f$ = Dir$(DataFilesPath & "EJTSClients????.dat")
Do Until f$ = ""
    If Len(f$) = 19 Then
        Y = Val(Mid$(f$, 12, 4))
        If Y > 0 Then
            lstDataFiles.AddItem Mid$(f$, 14, 2) & vbTab & f$ & vbTab & FileDateTime(DataFilesPath & f$)
            lstDataFiles.ItemData(lstDataFiles.NewIndex) = Y
            If Y = FileToOpen_Year Then
                lstDataFiles.ListIndex = lstDataFiles.NewIndex
                skipsel = True
            End If
        End If
    End If
    f$ = Dir$
Loop
If FileToOpen_Year > 0 And skipsel = False Then
    'This would happen if the command line says we should open a particular file, but we were unable to find it above
    FileToOpen_Year = 0
End If
If Not skipsel Then lstDataFiles.ListIndex = lstDataFiles.ListCount - 1
Exit Sub
e:
    lstDataFiles.AddItem "Error!"
    lstDataFiles.Enabled = False
End Sub

