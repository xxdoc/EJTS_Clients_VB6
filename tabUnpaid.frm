VERSION 5.00
Begin VB.Form tabUnpaid 
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
      Height          =   300
      IntegralHeight  =   0   'False
      ItemData        =   "tabUnpaid.frx":0000
      Left            =   0
      List            =   "tabUnpaid.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox lstResults 
      Height          =   540
      IntegralHeight  =   0   'False
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "tabUnpaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements ITab
Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

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
Const PROC_NAME = "tabUnpaid" & "." & "ITab_InitializeAfterDBLoad": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetTabStops lstResults.hwnd, 220, 320, 370

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
Const PROC_NAME = "tabUnpaid" & "." & "ITab_AfterTabShown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

Dim a&, cindex&
Dim RetTotFee&, RetTotOwe&, XChgTotFee&, XChgTotOwe&

lstResults.Clear

'Sort
lstSort.Clear
For cindex = 0 To ActiveDBInstance.Clients_Count - 1
    With ActiveDBInstance.Clients(cindex).c
        If .MoneyOwed <> NullLong Then
            lstSort.AddItem FormatClientName(fCustomListboxSorting, ActiveDBInstance.Clients(cindex).c)
            lstSort.ItemData(lstSort.NewIndex) = cindex
        End If
    End With
Next cindex
'Add to listbox
For a = 0 To lstSort.ListCount - 1
    cindex = lstSort.ItemData(a)
    With ActiveDBInstance.Clients(cindex).c
        If .MoneyOwed <> NullLong Then
            lstResults.AddItem FormatClientName(fUnpaid, ActiveDBInstance.Clients(cindex).c) & vbTab & _
                "Completed: " & FieldToString(.CompletionDate, mDateAsLongOrNULL) & vbTab & _
                "Fee: " & FieldToString(.PrepFee, mDollar) & vbTab & _
                "Owed: " & FieldToString(.MoneyOwed, mDollar)
            RetTotFee = RetTotFee + .PrepFee
            RetTotOwe = RetTotOwe + .MoneyOwed
        End If
    End With
Next a
lstResults.AddItem String$(185, 45)
lstResults.AddItem "Client totals" & vbTab & vbTab & "Total: " & FieldToString(RetTotFee, mDollar) & vbTab & "Total: " & FieldToString(RetTotOwe, mDollar)

lstResults.AddItem ""
lstResults.AddItem ""

'Sort
lstSort.Clear
For cindex = 0 To ActiveDBInstance.ExtraCharges_Count - 1
    With ActiveDBInstance.ExtraCharges(cindex)
        If .MoneyOwed <> NullLong Then
            lstSort.AddItem Format$(.CompletionDate, "yyyy-mm-dd") & .ClientName
            lstSort.ItemData(lstSort.NewIndex) = cindex
        End If
    End With
Next cindex
'Add to listbox
For a = 0 To lstSort.ListCount - 1
    cindex = lstSort.ItemData(a)
    With ActiveDBInstance.ExtraCharges(cindex)
        If .MoneyOwed <> NullLong Then
            lstResults.AddItem .ClientName & vbTab & _
                "Completed: " & FieldToString(.CompletionDate, mDateAsLongOrNULL) & vbTab & _
                "Fee: " & FieldToString(.PrepFee, mDollar) & vbTab & _
                "Owed: " & FieldToString(.MoneyOwed, mDollar)
            XChgTotFee = XChgTotFee + .PrepFee
            XChgTotOwe = XChgTotOwe + .MoneyOwed
        End If
    End With
Next a
lstResults.AddItem String$(185, 45)
lstResults.AddItem "Extra charge totals" & vbTab & vbTab & "Total: " & FieldToString(XChgTotFee, mDollar) & vbTab & "Total: " & FieldToString(XChgTotOwe, mDollar)

lstResults.AddItem ""
lstResults.AddItem ""

lstResults.AddItem String$(185, 45)
lstResults.AddItem "GRAND TOTAL" & vbTab & vbTab & "Total: " & FieldToString(RetTotFee + XChgTotFee, mDollar) & vbTab & "Total: " & FieldToString(RetTotOwe + XChgTotOwe, mDollar)

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
Const PROC_NAME = "tabUnpaid" & "." & "ITab_SetDefaultFocus": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

SetFocusWithoutErr lstResults

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
End Function

Private Function ITab_DestroyGDIObjects() As Boolean
End Function

Private Sub Form_Resize()
'errheader>
On Error Resume Next        'ALL ERRORS WILL BE IGNORED IN THIS PROCEDURE
'<errheader

lstResults.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'errheader>
Const PROC_NAME = "tabUnpaid" & "." & "Form_KeyDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabUnpaid" & "." & "Form_KeyUp": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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
Const PROC_NAME = "tabUnpaid" & "." & "Form_KeyPress": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
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

Private Sub lstResults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'errheader>
Const PROC_NAME = "tabUnpaid" & "." & "lstResults_MouseDown": Dim ERR_COUNT As Integer: On Error GoTo ERR_HANDLER
'<errheader

'Select item under mouse
Dim i&
i = SendMessage(lstResults.hwnd, LB_ITEMFROMPOINT, 0, (X / Screen.TwipsPerPixelX) + ((Y / Screen.TwipsPerPixelY) * &H10000))
If i > &HFFFF& Then
    lstResults.ListIndex = -1
Else
    i = (i And &HFFFF&)
    If Button = vbRightButton Then
        lstResults.ListIndex = i    'Listbox only does this for left click on a valid item
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
