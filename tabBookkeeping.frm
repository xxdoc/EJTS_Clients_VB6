VERSION 5.00
Begin VB.Form tabBookkeeping 
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
   Begin VB.PictureBox pctFrame 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   600
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Menu menName 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menNameRename 
         Caption         =   "&Rename..."
      End
      Begin VB.Menu menNameDelete 
         Caption         =   "&Delete item"
      End
      Begin VB.Menu menNameAdd 
         Caption         =   "&Add new item..."
         Index           =   0
      End
      Begin VB.Menu menNameAdd 
         Caption         =   "&Insert new item..."
         Index           =   1
      End
   End
   Begin VB.Menu menMonth 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menMonthEdit 
         Caption         =   "&Edit month data..."
      End
      Begin VB.Menu menMonthMarkPaid 
         Caption         =   "Mark &paid"
      End
   End
End
Attribute VB_Name = "tabBookkeeping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "tabBookkeeping"
Implements ITab

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private Const BOOK_OffsetX = 8
Private Const BOOK_OffsetY = 8
Private Const BOOK_NameWidth = 200
Private Const BOOK_ColumnWidth = 45
Private Const BOOK_TotalsWidth = 60
Private Const BOOK_LineHeight = 20
Private Const BOOK_OweColor = vbRed
Private BOOK_PartiallyPaidColor&
Private Const BOOK_BorderColor = vbBlack
Private Const BOOK_CellMargin = 2
Private Const BOOK_FontSize = 10
Private BOOK_Font&
Private BOOK_FontHeader&

Public BOOK_Clicked_BKIndex As Long
Public BOOK_Clicked_MonthIndex As Long
Public BOOK_Clicked_Changed As Boolean

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Standard
Private Function ITab_CreateGDIObjects() As Boolean
On Error GoTo ERR_HANDLER

BOOK_FontHeader = CreateFont2(pctFrame.hdc, "Arial", BOOK_FontSize, True, False, False, False)
BOOK_Font = CreateFont2(pctFrame.hdc, "Arial", BOOK_FontSize, False, False, False, False)

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_CreateGDIObjects", Err
End Function

'EHT=Standard
Private Function ITab_InitializeAfterDBLoad() As Boolean
On Error GoTo ERR_HANDLER


Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_InitializeAfterDBLoad", Err
End Function

'EHT=Standard
Private Sub ITab_AfterTabShown()
On Error GoTo ERR_HANDLER

Dim pcthdc&, tx&, ty&, a&, b&, vi&, t$
Dim monthtotals(11) As Long, owetotal As Long, rowtotal As Long, yeartotal As Long

pctFrame.Cls
pcthdc = pctFrame.hdc

'Vertical grid lines
ty = BOOK_OffsetY + ((ActiveDBInstance.Bookkeeping_Count + 2) * BOOK_LineHeight)
pctFrame.Line (BOOK_OffsetX, BOOK_OffsetY)-(BOOK_OffsetX, ty), BOOK_BorderColor
For b = 0 To 12
    tx = BOOK_OffsetX + BOOK_NameWidth + (b * BOOK_ColumnWidth)
    pctFrame.Line (tx, BOOK_OffsetY)-(tx, ty), BOOK_BorderColor
Next b
tx = BOOK_OffsetX + BOOK_NameWidth + (12 * BOOK_ColumnWidth) + BOOK_TotalsWidth
pctFrame.Line (tx, BOOK_OffsetY)-(tx, ty), BOOK_BorderColor
'Horizontal grid lines
tx = BOOK_OffsetX + BOOK_NameWidth + (BOOK_ColumnWidth * 12) + BOOK_TotalsWidth
For b = 0 To ActiveDBInstance.Bookkeeping_Count
    ty = BOOK_OffsetY + (b * BOOK_LineHeight) + BOOK_LineHeight
    pctFrame.Line (BOOK_OffsetX, ty)-(tx, ty), BOOK_BorderColor
Next b

'Header
SelectObject pcthdc, BOOK_FontHeader
SetTextColor pcthdc, vbBlack
SetTextAlign pcthdc, TA_LEFT
ty = BOOK_OffsetY + BOOK_CellMargin
t$ = "Name"
TextOut pcthdc, BOOK_OffsetX + BOOK_CellMargin, ty, t$, Len(t$)
For b = 1 To 12
    t$ = MonthName(b, True)
    SetTextAlign pcthdc, TA_RIGHT
    TextOut pcthdc, BOOK_OffsetX + BOOK_NameWidth + (b * BOOK_ColumnWidth) - BOOK_CellMargin, ty, t$, Len(t$)
Next b
t$ = "Total"
SetTextAlign pcthdc, TA_RIGHT
TextOut pcthdc, BOOK_OffsetX + BOOK_NameWidth + (12 * BOOK_ColumnWidth) + BOOK_TotalsWidth - BOOK_CellMargin, ty, t$, Len(t$)
vi = vi + 1

'Content
yeartotal = 0
BOOK_PartiallyPaidColor = RGB(210, 0, 255)
SelectObject pcthdc, BOOK_Font
For a = 0 To ActiveDBInstance.Bookkeeping_Count - 1
    ty = BOOK_OffsetY + (vi * BOOK_LineHeight) + BOOK_CellMargin
    t$ = ActiveDBInstance.Bookkeeping(a).DisplayName
    SelectObject pcthdc, BOOK_Font
    SetTextColor pcthdc, vbBlack
    SetTextAlign pcthdc, TA_LEFT
    TextOut pcthdc, BOOK_OffsetX + BOOK_CellMargin, ty, t$, Len(t$)
    rowtotal = 0
    For b = 0 To 11
        With ActiveDBInstance.Bookkeeping(a).Months(b)
            t$ = FieldToString(.PrepFee, mDollarOrNULL)
            If .PrepFee <> NullLong Then
                rowtotal = rowtotal + .PrepFee
                monthtotals(b) = monthtotals(b) + .PrepFee
                If .MoneyOwed = .PrepFee Then
                    owetotal = owetotal + .MoneyOwed
                    SetTextColor pcthdc, BOOK_OweColor
                ElseIf .MoneyOwed <> NullLong Then
                    owetotal = owetotal + .MoneyOwed
                    SetTextColor pcthdc, BOOK_PartiallyPaidColor
                Else
                    SetTextColor pcthdc, vbBlack
                End If
                SetTextAlign pcthdc, TA_RIGHT
                TextOut pcthdc, BOOK_OffsetX + BOOK_NameWidth + ((b + 1) * BOOK_ColumnWidth) - BOOK_CellMargin, ty, t$, Len(t$)
            End If
        End With
    Next b
    yeartotal = yeartotal + rowtotal
    t$ = FieldToString(rowtotal, mDollar)
    SelectObject pcthdc, BOOK_FontHeader
    SetTextColor pcthdc, vbBlack
    SetTextAlign pcthdc, TA_RIGHT
    TextOut pcthdc, BOOK_OffsetX + BOOK_NameWidth + (12 * BOOK_ColumnWidth) + BOOK_TotalsWidth - BOOK_CellMargin, ty, t$, Len(t$)
    vi = vi + 1
Next a

'Footer
SelectObject pcthdc, BOOK_FontHeader
SetTextColor pcthdc, vbBlack
SetTextAlign pcthdc, TA_LEFT
ty = BOOK_OffsetY + (vi * BOOK_LineHeight) + BOOK_CellMargin
t$ = "Totals"
TextOut pcthdc, BOOK_OffsetX + BOOK_CellMargin, ty, t$, Len(t$)
SetTextColor pcthdc, BOOK_OweColor
t$ = "(" & FieldToString(owetotal, mDollar) & " owed)"
TextOut pcthdc, BOOK_OffsetX + BOOK_CellMargin + 75, ty, t$, Len(t$)
SetTextColor pcthdc, vbBlack
For b = 0 To 11
    If monthtotals(b) = 0 Then
        t$ = ""
    Else
        t$ = FieldToString(monthtotals(b), mDollar)
    End If
    SetTextAlign pcthdc, TA_RIGHT
    TextOut pcthdc, BOOK_OffsetX + BOOK_NameWidth + ((b + 1) * BOOK_ColumnWidth) - BOOK_CellMargin, ty, t$, Len(t$)
Next b
t$ = FieldToString(yeartotal, mDollar)
SetTextColor pcthdc, vbBlack
SetTextAlign pcthdc, TA_RIGHT
TextOut pcthdc, BOOK_OffsetX + BOOK_NameWidth + (12 * BOOK_ColumnWidth) + BOOK_TotalsWidth - BOOK_CellMargin, ty, t$, Len(t$)

pctFrame.Refresh

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_AfterTabShown", Err
End Sub

'EHT=Standard
Private Sub ITab_SetDefaultFocus()
On Error GoTo ERR_HANDLER

SetFocusWithoutErr pctFrame

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

DeleteObject BOOK_FontHeader
DeleteObject BOOK_Font

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ITab_DestroyGDIObjects", Err
End Function

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

pctFrame.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
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
Private Sub menMonthEdit_Click()
On Error GoTo ERR_HANDLER

If Not menMonthEdit.Enabled Then Exit Sub

Dim frm As New frmBookkeepingEdit
If frm.Form_Show(BOOK_Clicked_BKIndex, BOOK_Clicked_MonthIndex) Then   'This will mark changed if necessary
    ITab_AfterTabShown
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menMonthEdit_Click", Err
End Sub

'EHT=Standard
Private Sub menMonthMarkPaid_Click()
On Error GoTo ERR_HANDLER

If Not menMonthMarkPaid.Enabled Then Exit Sub

With ActiveDBInstance.Bookkeeping(BOOK_Clicked_BKIndex).Months(BOOK_Clicked_MonthIndex)
    .MoneyOwed = NullLong
End With
tabLogFile.WriteLine "Bookkeeping entry marked paid: " & MonthName(BOOK_Clicked_MonthIndex + 1) & ", " & ActiveDBInstance.Bookkeeping(BOOK_Clicked_BKIndex).DisplayName
ITab_AfterTabShown
frmMain.SetChangedFlagAndIndication

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menMonthMarkPaid_Click", Err
End Sub

'EHT=Standard
Private Sub menNameRename_Click()
On Error GoTo ERR_HANDLER

If Not menNameRename.Enabled Then Exit Sub

Dim t$
t$ = InputBox("Edit name:", , ActiveDBInstance.Bookkeeping(BOOK_Clicked_BKIndex).DisplayName)
If t$ <> "" Then
    ActiveDBInstance.Bookkeeping(BOOK_Clicked_BKIndex).DisplayName = t$
    ITab_AfterTabShown
    frmMain.SetChangedFlagAndIndication
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menNameRename_Click", Err
End Sub

'EHT=Standard
Private Sub menNameDelete_Click()
On Error GoTo ERR_HANDLER

If Not menNameDelete.Enabled Then Exit Sub

If MsgBox("Are you sure you want to delete '" & ActiveDBInstance.Bookkeeping(BOOK_Clicked_BKIndex).DisplayName & "'?", vbQuestion Or vbYesNo) = vbYes Then
    DB_RemoveBookkeepingJob ActiveDBInstance, BOOK_Clicked_BKIndex
    ITab_AfterTabShown
    frmMain.SetChangedFlagAndIndication
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menNameDelete_Click", Err
End Sub

'EHT=Standard
Private Sub menNameAdd_Click(Index As Integer)
On Error GoTo ERR_HANDLER

If Not menNameAdd(Index).Enabled Then Exit Sub

Dim bk As BookkeepingJob, bkindex&, t$, BeforeIndex&
For bkindex = 0 To UBound(bk.Months)
    With bk.Months(bkindex)
        .CompletionDate = NullLong
        .PrepFee = NullLong
        .MoneyOwed = NullLong
    End With
Next bkindex
If Index = 0 Then BeforeIndex = -1 Else BeforeIndex = BOOK_Clicked_BKIndex
bkindex = DB_AddBookkeepingJob(ActiveDBInstance, bk, BeforeIndex)
ITab_AfterTabShown

t$ = InputBox("Enter new name:")
If t$ <> "" Then
    ActiveDBInstance.Bookkeeping(bkindex).DisplayName = t$
    frmMain.SetChangedFlagAndIndication
Else
    DB_RemoveBookkeepingJob ActiveDBInstance, bkindex
End If
ITab_AfterTabShown

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "menNameAdd_Click", Err
End Sub

'EHT=Standard
Private Sub pctFrame_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ERR_HANDLER

Dim bkindex&, cellindex&

If BOOK_GetCellFromXY(X, Y, bkindex&, cellindex&) Then
    If cellindex < 0 Then
        BOOK_Clicked_BKIndex = bkindex
        If Button = vbLeftButton Then
        ElseIf Button = vbRightButton Then
            menNameRename.Enabled = True And (ActiveDBInstance.IsWriteable)
            menNameDelete.Enabled = True And (ActiveDBInstance.IsWriteable)
            menNameAdd(0).Enabled = True And (ActiveDBInstance.IsWriteable)
            menNameAdd(1).Enabled = True And (ActiveDBInstance.IsWriteable)
            PopupMenu menName   'No With blocks!!!
        End If
    Else
        BOOK_Clicked_BKIndex = bkindex
        BOOK_Clicked_MonthIndex = cellindex
        If Button = vbLeftButton Then
            menMonthEdit_Click
        ElseIf Button = vbRightButton Then
            menMonthMarkPaid.Enabled = (ActiveDBInstance.Bookkeeping(bkindex).Months(cellindex).MoneyOwed <> NullLong) And (ActiveDBInstance.IsWriteable)
            PopupMenu menMonth, , , , menMonthEdit  'No With blocks!!!
        End If
    End If
Else
    If Button = vbRightButton Then
        'Outside table
        menNameRename.Enabled = False And (ActiveDBInstance.IsWriteable)
        menNameDelete.Enabled = False And (ActiveDBInstance.IsWriteable)
        menNameAdd(0).Enabled = True And (ActiveDBInstance.IsWriteable)
        menNameAdd(1).Enabled = True And (ActiveDBInstance.IsWriteable)
        PopupMenu menName   'No With blocks!!!
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "pctFrame_MouseDown", Err
End Sub

'EHT=Standard
Function BOOK_GetCellFromXY(ByVal X&, ByVal Y&, ByRef bkindex&, ByRef cellindex&) As Boolean
On Error GoTo ERR_HANDLER

X = X - BOOK_OffsetX
Y = Y - BOOK_OffsetY

bkindex = Int(Y / BOOK_LineHeight) - 1
If bkindex < 0 Then Exit Function
If bkindex >= ActiveDBInstance.Bookkeeping_Count Then Exit Function

If X < 0 Then Exit Function
If X < BOOK_NameWidth Then
    cellindex = -1
Else
    cellindex = Int((X - BOOK_NameWidth) / BOOK_ColumnWidth)
    If cellindex > 11 Then Exit Function
End If
BOOK_GetCellFromXY = True

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "BOOK_GetCellFromXY", Err
End Function

