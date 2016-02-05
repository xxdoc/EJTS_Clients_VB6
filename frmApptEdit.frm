VERSION 5.00
Begin VB.Form frmApptEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Appointment"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApptEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   361
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton btnMoveDown 
      Caption         =   "Move &Down"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton btnMoveUp 
      Caption         =   "Move &Up"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ListBox lstClients 
      Height          =   1335
      IntegralHeight  =   0   'False
      ItemData        =   "frmApptEdit.frx":000C
      Left            =   120
      List            =   "frmApptEdit.frx":000E
      TabIndex        =   7
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox txtField 
      Height          =   1215
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "50"
      Top             =   1080
      Width           =   3255
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
      Left            =   4200
      TabIndex        =   3
      Tag             =   "13"
      Top             =   360
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
      Left            =   1440
      TabIndex        =   1
      Tag             =   "40"
      Top             =   360
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
      Left            =   2880
      TabIndex        =   2
      Tag             =   "13"
      Top             =   360
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "30"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
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
      TabIndex        =   5
      Top             =   4440
      Width           =   3135
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
      TabIndex        =   21
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   352
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Flags:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblFlag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Appt Didn't Happen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   19
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblFlag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Called"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   18
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblFlag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reminder Call"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   17
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Day:"
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
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Time:"
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
      Left            =   1440
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Time slot:"
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
      Left            =   2880
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "# of slots:"
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
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Notes:"
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
      TabIndex        =   16
      Top             =   840
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   3480
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmApptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "frmApptEdit"

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again
Public TabOrderSetting As String            'This is set in Form_Show

Private Enum FieldName
    fDay
    fActualTime
    fTimeSlot
    fNumSlots
    fNotes
End Enum

Private ClientsToDelete&()
Private ClientsToDelete_Count&
Private ClientsToAdd&()
Private ClientsToAdd_Count&

Private thisID&
Private Changed As Boolean

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Function Form_Show(aID&, Optional FocusNotesBoxFirst As Boolean = False) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

Dim aindex&, a&, i&

thisID = aID
aindex = DB_FindAppointmentIndex(ActiveDBInstance, aID)
If aindex < 0 Then Err.Raise 1, , "Appointment #" & aID & " not found!"

'Set the tab order
TabOrderSetting = "GLOBAL_TabOrder_ApptEdit"
SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)

With ActiveDBInstance.Appointments(aindex)
    Me.Caption = Me.Caption & " (#" & .ID & ")"

    FieldToTextbox txtField(fNumSlots), .NumSlots

    FieldToTextbox txtField(fTimeSlot), .ApptTimeSlot
    'FieldToTextbox txtField(fTimeSlot), Appointment_FirstSlotTime + (.ApptTimeSlot * Appointment_SlotLength)

    FieldToTextbox txtField(fDay), .ApptDate

    FieldToTextbox txtField(fActualTime), .ApptActualTime

    FieldToTextbox txtField(fNotes), .Notes

    For a = 0 To AppointmentFlags_DATAITEMUBOUND
        If Flag_IsSet(.Flags, 2 ^ a) Then
            lblFlag(a).BackStyle = 1
            lblFlag(a).BorderStyle = 1
        Else
            lblFlag(a).BackStyle = 0
            lblFlag(a).BorderStyle = 0
        End If
    Next a

    lstClients.Clear
    For a = 0 To .ClientID_Count - 1
        i = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(a))
        lstClients.AddItem FormatClientName(fSearchResults, ActiveDBInstance.Clients(i).c)
        lstClients.ItemData(a) = .ClientIDs(a)
    Next a
End With

lblChangeTabOrder.Move Me.ScaleWidth - lblChangeTabOrder.Width - 1, Me.ScaleHeight - lblChangeTabOrder.Height - 1

lstClients_Click
btnSave.Enabled = ActiveDBInstance.IsWriteable
frmMain.IdlePauseTimeout
If FocusNotesBoxFirst Then txtField(fNotes).TabIndex = 0
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

Dim tempappt As Appointment, aindex&, a&, cindex&
aindex = DB_FindAppointmentIndex(ActiveDBInstance, thisID)
If aindex < 0 Then Err.Raise 1, , "Unable to save. Appointment #" & thisID & " not found!"

tempappt = ActiveDBInstance.Appointments(aindex)   'Make a temp copy
With tempappt
    FieldFromTextbox txtField(fNumSlots), .NumSlots
    If .NumSlots < 1 Then
        ShowErrorMsg "Num slots is invalid"
        SetFocusWithoutErr txtField(fNumSlots)
        Exit Sub
    End If

    FieldFromTextbox txtField(fTimeSlot), .ApptTimeSlot
    '.ApptTimeSlot = (CDate(txtField(fTimeSlot)) - Appointment_FirstSlotTime) / Appointment_SlotLength
'    If <outside timeslot range> Then
'        ShowErrorMsg "Time slot is invalid"
'        SetFocusWithoutErr txtField(fTimeSlot)
'        Exit Sub
'    End If

    FieldFromTextbox txtField(fDay), .ApptDate
'    If <outside bitmap range> Then
'        ShowErrorMsg "Day is invalid"
'        SetFocusWithoutErr txtField(fDay)
'        Exit Sub
'    End If

    FieldFromTextbox txtField(fActualTime), .ApptActualTime

    FieldFromTextbox txtField(fNotes), .Notes

    .Flags = 0
    For a = 0 To AppointmentFlags_DATAITEMUBOUND
        If lblFlag(a).BackStyle = 1 Then
            .Flags = .Flags Or (2 ^ a)
        End If
    Next a

    .ClientID_Count = lstClients.ListCount
    If .ClientID_Count = 0 Then
        Erase .ClientIDs
    Else
        ReDim .ClientIDs(.ClientID_Count - 1)
        For a = 0 To .ClientID_Count - 1
            .ClientIDs(a) = lstClients.ItemData(a)
        Next a
    End If

    'Update bitmap
    If DB_SlotsIsAvail(ActiveDBInstance, .ApptDate, .ApptTimeSlot, .NumSlots, .ID) Then
        'Clear old position
        DB_SlotsClear ActiveDBInstance, ActiveDBInstance.Appointments(aindex).ApptDate, ActiveDBInstance.Appointments(aindex).ApptTimeSlot, ActiveDBInstance.Appointments(aindex).NumSlots
        'Put appointment into new position
        DB_SlotsFill ActiveDBInstance, .ApptDate, .ApptTimeSlot, .NumSlots, aindex
    Else
        ShowErrorMsg "Selected appointment slots are not available!"
        Exit Sub
    End If

    'Write temp copy back to database
    ActiveDBInstance.Appointments(aindex) = tempappt

    'Regenerate Temp data
    For a = 0 To .ClientID_Count - 1
        cindex = DB_FindClientIndex(ActiveDBInstance, .ClientIDs(a))
        ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
    Next a

    'Add OpNotes
    For a = 0 To ClientsToDelete_Count - 1
        cindex = DB_FindClientIndex(ActiveDBInstance, ClientsToDelete(a))
        AddOpNote ActiveDBInstance.Clients(cindex).c.OpNotes, "Cancelled appt: " & FormatApptTime$(.ApptDate, .ApptActualTime)
        ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData = True
    Next a
    For a = 0 To ClientsToAdd_Count - 1
        cindex = DB_FindClientIndex(ActiveDBInstance, ClientsToAdd(a))
        AddOpNote ActiveDBInstance.Clients(cindex).c.OpNotes, "Scheduled appt: " & FormatApptTime$(.ApptDate, .ApptActualTime)
    Next a

    frmMain.DayTotal_Update
    frmMain.SetChangedFlagAndIndication
    tabLogFile.WriteLine "Edited " & DB_FormatApptClientList(ActiveDBInstance, tempappt) & ": " & FormatApptTime$(.ApptDate, .ApptActualTime) & ", " & FormatNumApptSlots(.NumSlots)
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
Private Sub btnAdd_Click()
On Error GoTo ERR_HANDLER

If Not btnAdd.Enabled Then Exit Sub

SetFocusWithoutErr lstClients

Dim newID&, NewIndex&, a&, b&, f As Boolean, t$

TryAgain:
t$ = InputBox("Enter client ID#", , t$)
If t$ = "" Then Exit Sub
If Not ConvertToLong(t$, newID) Then
    ShowErrorMsg "You must enter a numeric value"
    GoTo TryAgain
End If

NewIndex = DB_FindClientIndex(ActiveDBInstance, newID)
If NewIndex < 0 Then
    ShowErrorMsg "Cannot find client ID#" & newID
    GoTo TryAgain
End If
With ActiveDBInstance.Clients(NewIndex).c
    'Add to listbox
    lstClients.AddItem FormatClientName(fSearchResults, ActiveDBInstance.Clients(NewIndex).c)
    lstClients.ItemData(lstClients.NewIndex) = newID

    'Search for newID in ClientsToDelete
    For a = 0 To ClientsToDelete_Count - 1
        If ClientsToDelete(a) = newID Then
            f = True
            Exit For
        End If
    Next a
    If f Then
        'Found, remove it (don't add to ClientsToAdd, since it is already in the original)
        For b = a To ClientsToDelete_Count - 2
            ClientsToDelete(b) = ClientsToDelete(b + 1)
        Next b
        ClientsToDelete_Count = ClientsToDelete_Count - 1
        If ClientsToDelete_Count = 0 Then
            Erase ClientsToDelete
        Else
            ReDim Preserve ClientsToDelete(ClientsToDelete_Count - 1)
        End If
    Else
        'Not found, add to ClientsToAdd
        ReDim Preserve ClientsToAdd(ClientsToAdd_Count)
        ClientsToAdd(ClientsToAdd_Count) = newID
        ClientsToAdd_Count = ClientsToAdd_Count + 1
    End If

    lstClients_Click
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnAdd_Click", Err
End Sub

'EHT=Standard
Private Sub btnMoveDown_Click()
On Error GoTo ERR_HANDLER

If Not btnMoveDown.Enabled Then Exit Sub

Dim i1&, i2&, td&, ts$
SetFocusWithoutErr lstClients
i1 = lstClients.ListIndex
If (i1 >= 0) And (i1 < (lstClients.ListCount - 1)) Then
    i2 = i1 + 1
    td = lstClients.ItemData(i2)
    ts$ = lstClients.List(i2)
    lstClients.ItemData(i2) = lstClients.ItemData(i1)
    lstClients.List(i2) = lstClients.List(i1)
    lstClients.ItemData(i1) = td
    lstClients.List(i1) = ts$
    lstClients.ListIndex = i2
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnMoveDown_Click", Err
End Sub

'EHT=Standard
Private Sub btnMoveUp_Click()
On Error GoTo ERR_HANDLER

If Not btnMoveUp.Enabled Then Exit Sub

Dim i1&, i2&, td&, ts$
SetFocusWithoutErr lstClients
i1 = lstClients.ListIndex
If i1 > 0 Then
    i2 = i1 - 1
    td = lstClients.ItemData(i2)
    ts$ = lstClients.List(i2)
    lstClients.ItemData(i2) = lstClients.ItemData(i1)
    lstClients.List(i2) = lstClients.List(i1)
    lstClients.ItemData(i1) = td
    lstClients.List(i1) = ts$
    lstClients.ListIndex = i2
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnMoveUp_Click", Err
End Sub

'EHT=Standard
Private Sub btnRemove_Click()
On Error GoTo ERR_HANDLER

If Not btnRemove.Enabled Then Exit Sub

Dim i&, cID&, a&, b&, f As Boolean
SetFocusWithoutErr lstClients
i = lstClients.ListIndex
If (i >= 0) Then
    cID = lstClients.ItemData(i)

    'Search for newID in ClientsToAdd
    For a = 0 To ClientsToAdd_Count - 1
        If ClientsToAdd(a) = cID Then
            f = True
            Exit For
        End If
    Next a
    If f Then
        'Found, remove it (don't add to ClientsToDelete, since it was never in the original)
        For b = a To ClientsToAdd_Count - 2
            ClientsToAdd(b) = ClientsToAdd(b + 1)
        Next b
        ClientsToAdd_Count = ClientsToAdd_Count - 1
        If ClientsToAdd_Count = 0 Then
            ReDim ClientsToAdd(0)
        Else
            ReDim Preserve ClientsToAdd(ClientsToAdd_Count - 1)
        End If
    Else
        'Not found, add to ClientsToDelete
        ReDim Preserve ClientsToDelete(ClientsToDelete_Count)
        ClientsToDelete(ClientsToDelete_Count) = cID
        ClientsToDelete_Count = ClientsToDelete_Count + 1
    End If

    lstClients.RemoveItem i
    lstClients_Click
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnRemove_Click", Err
End Sub

'EHT=Standard
Private Sub lstClients_Click()
On Error GoTo ERR_HANDLER

Dim i&

i = lstClients.ListIndex
btnMoveUp.Enabled = (i > 0) And (ActiveDBInstance.IsWriteable)
btnMoveDown.Enabled = (i >= 0) And (i < (lstClients.ListCount - 1)) And (ActiveDBInstance.IsWriteable)
btnAdd.Enabled = (ActiveDBInstance.IsWriteable)
btnRemove.Enabled = (i >= 0) And (ActiveDBInstance.IsWriteable)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstClients_Click", Err
End Sub

'EHT=Standard
Private Sub lstClients_DblClick()
On Error GoTo ERR_HANDLER

Dim i&, c As CClient
i = lstClients.ListIndex
If i >= 0 Then
    Set c = frmMain.NEWDATABASE.Client(lstClients.ItemData(i))
    If Not c Is Nothing Then
        Dim frm As New frmClientEditPost
        frm.Form_Show fEdit, c, , Me     'This will mark changed if necessary
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstClients_DblClick", Err
End Sub

'EHT=Standard
Private Sub lblFlag_Click(Index As Integer)
On Error GoTo ERR_HANDLER

With lblFlag(Index)
    If .BackStyle = 1 Then
        .BackStyle = 0
        .BorderStyle = 0
    Else
        .BackStyle = 1
        .BorderStyle = 1
    End If
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lblFlag_Click", Err
End Sub

'EHT=Standard
Private Sub txtField_LostFocus(Index As Integer)
On Error GoTo ERR_HANDLER

'If Index = fTimeSlot Then
'    'Custom handler to round to nearest slot time
'    Dim ts&
'    If IsDate(txtField(fTimeSlot).Text) Then
'        ts = (CDate(txtField(fTimeSlot).Text) - Appointment_FirstSlotTime) / Appointment_SlotLength
'        If (ts < 0) Or (ts > Appointment_NumSlotsUB) Then
'            txtField(fTimeSlot).Text = ""
'        Else
'            FieldToTextbox txtField(fTimeSlot), Appointment_FirstSlotTime + (ts * Appointment_SlotLength)
'        End If
'    Else
'        txtField(fTimeSlot).Text = ""
'    End If
'Else
    LostFocusFormat txtField(Index)
'End If

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

