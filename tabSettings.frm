VERSION 5.00
Begin VB.Form tabSettings 
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
      Height          =   735
      IntegralHeight  =   0   'False
      Left            =   10320
      Sorted          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pctDebugActions 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   12015
      TabIndex        =   5
      Top             =   4320
      Width           =   12015
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "listbyage"
         Height          =   495
         Index           =   6
         Left            =   7800
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "removexcell"
         Height          =   495
         Index           =   5
         Left            =   6360
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "fixnamecase"
         Height          =   495
         Index           =   4
         Left            =   4920
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "copynames"
         Height          =   495
         Index           =   3
         Left            =   3480
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "ranges"
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "dateofdeath"
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton btnDebugCode 
         Caption         =   "copyfees"
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Debug Actions (each button will show a confirmation message first, explaining what it does):"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   9375
      End
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFE1C4&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstSettings 
      Height          =   375
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   7080
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   5895
   End
   Begin VB.ListBox lstSettings 
      Height          =   3135
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Settings Specific to _________ database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Global Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "tabSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabSetting"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private LastListIndex(1) As Integer
Private LastTopIndex(1) As Integer
Private EditingList As Integer          '0 is Global, 1 is Specific
Private EditingListIndex As Integer     'Index in the on-screen listbox
Private EditingSettingName As String
Private TextboxLeftOffset As Long

'EHT=Standard
Private Sub btnDebugCode_Click(Index As Integer)
On Error GoTo ERR_HANDLER

If Not ActiveDBInstance.IsWriteable Then
    ShowErrorMsg "Not available in read-only mode!"
    Exit Sub
End If

Dim a&, b&, t$

Const WILL_CHANGE = vbCrLf & vbCrLf & "This WILL CHANGE the database. Cancel if you do not know what this function is for."
Const WILL_NOT_CHANGE = vbCrLf & vbCrLf & "The database will not be modified."

Select Case btnDebugCode(Index).Caption
Case "copyfees"
    If MsgBox("This will copy to the clipboard a sorted list of all clients with the CompletedReturn flag set, along with their last year's prep fees, current year's prep fees, and amount owed." & WILL_NOT_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            If Flag_IsSet(.c.Flags, CompletedReturn) Then
                lstSort.AddItem FormatClientName(fPullFiles, .c)
                lstSort.ItemData(lstSort.NewIndex) = a
            End If
        End With
    Next a
    t$ = "ID" & vbTab & "Name" & vbTab & "LYFee" & vbTab & "CYFee" & vbTab & "CYOwed" & vbCrLf
    For a = 0 To lstSort.ListCount - 1
        With ActiveDBInstance.Clients(lstSort.ItemData(a))
            t$ = t$ & .c.ID & vbTab & FormatClientName(fSearchResults, .c) & vbTab & FieldToString(.c.LastYear_PrepFee, mDollarOrNULL) & vbTab & FieldToString(.c.PrepFee, mDollarOrNULL) & vbTab & FieldToString(.c.MoneyOwed, mDollarOrNULL) & vbCrLf
        End With
    Next a
    Clipboard.Clear
    Clipboard.SetText t$
    MsgBox "Data copied to the clipboard.", vbInformation

Case "dateofdeath"
    If MsgBox("This will copy to the clipboard a sorted list of all client entries where one or both of the people have died, along with the dates of death." & WILL_NOT_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            lstSort.AddItem FormatClientName(fPullFiles, .c)
            lstSort.ItemData(lstSort.NewIndex) = a
        End With
    Next a
    For a = 0 To lstSort.ListCount - 1
        With ActiveDBInstance.Clients(lstSort.ItemData(a))
            If (.c.Person1.DOD <> NullLong) Or (.c.Person2.DOD <> NullLong) Then
                t$ = t$ & .c.ID & vbTab & .c.Person1.Last & vbTab & .c.Person1.First & vbTab & .c.Person2.First & vbTab & IIf(.c.Person1.DOD = NullLong, "", "D") & vbTab & IIf(.c.Person2.DOD = NullLong, "", "D") & vbTab & FormatClientName(fSearchResults, .c) & vbCrLf
            End If
        End With
    Next a
    Clipboard.Clear
    Clipboard.SetText t$

Case "ranges"
    If MsgBox("This loads a text file named ""newest oldest.txt"" and initializes each of the mentioned clients with the OldestYearFiled and NewestYearFiled information." & WILL_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    Dim fh As CMNMOD_CFileHandler
    Dim l$()
    Dim cindex&
    Set fh = OpenFile("newest oldest.txt", mLineByLine_Input)
    Do Until fh.EndOfFile
        t$ = fh.ReadLine
        l$ = Split(t$, vbTab)
        a = Val(l$(0))
        cindex = DB_FindClientIndex(ActiveDBInstance, a)
        If cindex >= 0 Then
            With ActiveDBInstance.Clients(cindex).c
                a = Val(l$(1))
                If a = 9900 Then a = NullLong
                If .OldestYearFiled = 9900 Then .OldestYearFiled = NullLong
                If .OldestYearFiled <> NullLong Then
                    If .OldestYearFiled <> a Then Stop
                End If
                .OldestYearFiled = a

                a = Val(l$(2))
                If a = 9900 Then a = NullLong
                If .NewestYearFiled = 9900 Then .NewestYearFiled = NullLong
                If .NewestYearFiled <> NullLong Then
                    If .NewestYearFiled <> a Then Stop
                End If
                .NewestYearFiled = a
            End With
        Else
            Stop
        End If
    Loop
    fh.CloseFile
    Stop

Case "copynames"
    If MsgBox("This will copy to the clipboard a sorted list of all client entries, along with their OldestYearFiled and NewestYearFiled information." & WILL_NOT_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            lstSort.AddItem FormatClientName(fPullFiles, .c)
            lstSort.ItemData(lstSort.NewIndex) = a
        End With
    Next a
    For a = 0 To lstSort.ListCount - 1
        With ActiveDBInstance.Clients(lstSort.ItemData(a))
            t$ = t$ & .c.ID & vbTab & FormatClientName(fPullFiles, .c) & vbTab & FieldToString(.c.OldestYearFiled, mYearOrNULL) & vbTab & FieldToString(.c.NewestYearFiled, mYearOrNULL) & vbCrLf
        End With
    Next a
    Clipboard.Clear
    Clipboard.SetText t$

Case "fixnamecase"
    If MsgBox("This will change every email in the database to lowercase." & WILL_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            .c.Person1.Email = LCase(.c.Person1.Email)
            .c.Person2.Email = LCase(.c.Person2.Email)
        End With
    Next a
    frmMain.SetChangedFlagAndIndication

Case "removexcell"
    If MsgBox("This will remove any xCELL suffix from every phone number in the database, while leaving other extensions (x1234)." & WILL_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            If LCase$(Right$(.c.Person1.Phone, 5)) = "xcell" Then
                .c.Person1.Phone = Left$(.c.Person1.Phone, Len(.c.Person1.Phone) - 5)
                b = b + 1
            End If
            If LCase$(Right$(.c.Person2.Phone, 5)) = "xcell" Then
                .c.Person2.Phone = Left$(.c.Person2.Phone, Len(.c.Person2.Phone) - 5)
                b = b + 1
            End If
        End With
    Next a
    frmMain.SetChangedFlagAndIndication
    MsgBox b & " total xCELL extensions removed."

Case "listbyage"
    If MsgBox("Creates a ""listbyage.csv"" with a list of all living individuals, sorted by current age." & WILL_NOT_CHANGE, vbQuestion Or vbOKCancel) = vbCancel Then Exit Sub
    
    Dim tod As Long, pers As PersonStruct
    tod = Date
    Set fh = OpenFile(AppPath & "listbyage.csv", mLineByLine_Output)
    fh.WriteLine "Client #,Individual,Date of Birth,Age as of " & Format(tod, "m/d/yyyy")
    lstSort.Clear
    For a = 0 To ActiveDBInstance.Clients_Count - 1
        With ActiveDBInstance.Clients(a)
            If (Len(.c.Person1.First) > 0) And (.c.Person1.DOB <> NullLong) And (.c.Person1.DOD = NullLong) Then
                lstSort.AddItem Format(.c.Person1.DOB, "0000000000") & vbTab & .c.Person1.Last & ", " & .c.Person1.First
                lstSort.ItemData(lstSort.NewIndex) = a
            End If
            If (Len(.c.Person2.First) > 0) And (.c.Person2.DOB <> NullLong) And (.c.Person2.DOD = NullLong) Then
                t$ = .c.Person2.Last
                If Len(t$) = 0 Then t$ = .c.Person1.Last
                lstSort.AddItem Format(.c.Person2.DOB, "0000000000") & vbTab & t$ & ", " & .c.Person2.First
                lstSort.ItemData(lstSort.NewIndex) = -a
            End If
        End With
    Next a
    For a = 0 To lstSort.ListCount - 1
        b = lstSort.ItemData(a)
        If b >= 0 Then
            pers = ActiveDBInstance.Clients(b).c.Person1
            t$ = pers.Last
        Else
            b = -b
            pers = ActiveDBInstance.Clients(b).c.Person2
            If Len(pers.Last) > 0 Then
                t$ = pers.Last
            Else
                t$ = ActiveDBInstance.Clients(b).c.Person1.Last
            End If
        End If
        t$ = t$ & ", " & pers.First
        If Len(pers.Initial) > 0 Then t$ = t$ & " " & pers.Initial
        fh.WriteLine b & ",""" & t$ & """," & Format(pers.DOB, "m/d/yyyy") & "," & CalculateAge(pers.DOB, tod)
    Next a
    fh.CloseFile
    ShowInfoMsg "A list of all living individuals, sorted by current age, has been saved to """ & fh.FullPath & """"

End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnDebugCode_Click", Err
End Sub

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

SetTabStops lstSettings(0).hwnd, 150, 180
SetTabStops lstSettings(1).hwnd, 150, 180
'Pixel position of second tabstop (trial & error, since tabstops aren't defined by pixels)
TextboxLeftOffset = 317 '= 180
lblTitle(1).Caption = "Settings Specific to " & FileToOpen_Year & " database"

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

Dim a&

EditingList = -1
EditingListIndex = -1
txtEdit.Visible = False

'Save positions
For a = 0 To 1
    LastTopIndex(a) = lstSettings(a).TopIndex
    LastListIndex(a) = lstSettings(a).ListIndex
Next a

'Global
lstSettings(0).Clear
For a = 0 To GlobalSettings_Count - 1
    If Left$(GlobalSettings(a).sName, 1) <> "_" Then 'Hide any item that begins with '_'
        lstSettings(0).AddItem SettingToListboxLine(GlobalSettings(a))
        lstSettings(0).ItemData(lstSettings(0).NewIndex) = a
    End If
Next a

'Specific
lstSettings(1).Clear
For a = 0 To ActiveDBInstance.Settings_Count - 1
    If Left$(ActiveDBInstance.Settings(a).sName, 1) <> "_" Then 'Hide any item that begins with '_'
        lstSettings(1).AddItem SettingToListboxLine(ActiveDBInstance.Settings(a))
        lstSettings(1).ItemData(lstSettings(1).NewIndex) = a
    End If
Next a

'Restore positions
For a = 0 To 1
    lstSettings(a).TopIndex = LastTopIndex(a)
    lstSettings(a).ListIndex = LastListIndex(a)
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

If txtEdit.Visible Then
    SetFocusWithoutErr txtEdit
Else
    SetFocusWithoutErr lstSettings(1)
End If

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

Dim l&, t&, w&, h&

l = 0
w = (Me.ScaleWidth - 16) / 2
t = lblTitle(0).Height + 8
h = Me.ScaleHeight - t - pctDebugActions.Height - 8

pctDebugActions.Move 0, Me.ScaleHeight - pctDebugActions.Height, w

lblTitle(0).Move l, 0, w
lstSettings(0).Move l, t, w, h
l = w + 8
lblTitle(1).Move l, 0, w
lstSettings(1).Move l, t, w, h
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
Private Sub lstSettings_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER

'Select item under mouse
Dim i&
i = SendMessage(lstSettings(Index).hwnd, LB_ITEMFROMPOINT, 0, (X / Screen.TwipsPerPixelX) + ((Y / Screen.TwipsPerPixelY) * &H10000))
If i > &HFFFF& Then
    lstSettings(Index).ListIndex = -1
Else
    i = (i And &HFFFF&)
    If Button = vbRightButton Then
        lstSettings(Index).ListIndex = i    'Listbox only does this for left click on a valid item
    End If
    lstSettings_Click Index    'Without this, the _Click event doesn't happen until MouseUp. Although, calling it here results in it
                               'being called twice for every click, at least the textbox gets moved the moment the selection moves
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSettings_MouseDown", Err
End Sub

'EHT=Standard
Private Sub lstSettings_Click(Index As Integer)
On Error GoTo ERR_HANDLER

Dim i%, ss&, r As RECT, otherindex%
i = lstSettings(Index).ListIndex
If i < 0 Then
    'Hide the textbox
    EditingList = -1
    EditingListIndex = -1
    EditingSettingName = ""
    txtEdit.Visible = False
Else
    'Hide the selection in the other list
    otherindex = (Not (Index - 1)) + 1
    lstSettings(otherindex).ListIndex = -1      'This may call lstSettings_Click

    If ActiveDBInstance.IsWriteable Then
        'Show the edit box
        EditingList = Index
        EditingListIndex = -1

        'Fill the edit box
        EditingSettingName = lstSettings(EditingList).List(i)
        EditingSettingName = Left$(EditingSettingName, InStr(EditingSettingName, vbTab) - 1)
        ss = txtEdit.SelStart   'Save cursor position
        txtEdit.Text = DB_GetSetting(ActiveDBInstance, EditingSettingName, True)
        txtEdit.SelStart = ss

        'Position it and show it
        SendMessageByRect lstSettings(Index).hwnd, LB_GETITEMRECT, i, r
        txtEdit.Move lstSettings(Index).Left + r.Left + TextboxLeftOffset, lstSettings(Index).Top + r.Top + 2, r.Right - TextboxLeftOffset + 2
        txtEdit.Visible = True
        SetFocusWithoutErr txtEdit
        EditingListIndex = i        'Leave this for last, because txtEdit_Change uses it
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstSettings_Click", Err
End Sub

'EHT=Standard
Private Sub txtEdit_Change()
On Error GoTo ERR_HANDLER

Dim t$
If ActiveDBInstance.IsWriteable Then
    If EditingList < 0 Then Exit Sub
    If EditingListIndex < 0 Then Exit Sub
    t$ = txtEdit.Text
    If LCase$(t$) = "(default)" Then
        'Hidden feature: setting any value to "(default)" will reset that setting to its default value
        If Not DB_SetDefaultSettingValue(ActiveDBInstance, EditingSettingName) Then
            'But if there is not default value stored in this code, set it to the actual text "(default)"
            DB_SetSetting ActiveDBInstance, EditingSettingName, t$
        End If
    Else
        DB_SetSetting ActiveDBInstance, EditingSettingName, t$
    End If
    If EditingList = 0 Then
        lstSettings(EditingList).List(EditingListIndex) = SettingToListboxLine(GlobalSettings(lstSettings(EditingList).ItemData(EditingListIndex)))
    ElseIf EditingList = 1 Then
        lstSettings(EditingList).List(EditingListIndex) = SettingToListboxLine(ActiveDBInstance.Settings(lstSettings(EditingList).ItemData(EditingListIndex)))
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtEdit_Change", Err
End Sub

'EHT=Standard
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

Dim i%
Select Case KeyCode
Case vbKeyUp
    i = lstSettings(EditingList).ListIndex
    If i > 0 Then lstSettings(EditingList).ListIndex = i - 1
    KeyCode = 0
Case vbKeyDown
    i = lstSettings(EditingList).ListIndex
    If i < lstSettings(EditingList).ListCount - 1 Then lstSettings(EditingList).ListIndex = i + 1
    KeyCode = 0
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtEdit_KeyDown", Err
End Sub

'EHT=Standard
Private Function SettingToListboxLine(s As Setting) As String
On Error GoTo ERR_HANDLER

SettingToListboxLine = s.sName & vbTab & Choose(s.sType + 1, "Str", "Num", "Date", "T/F", "Dbl") & vbTab & DB_FormatSettingForScreen(s)

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SettingToListboxLine", Err
End Function

