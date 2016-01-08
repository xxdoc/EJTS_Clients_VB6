VERSION 5.00
Begin VB.UserControl CustomListbox 
   BackColor       =   &H80000005&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   191
End
Attribute VB_Name = "CustomListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "vbaccelerator Owner Draw Combo and List box control."
Option Explicit
Private Const MOD_NAME = "CustomListbox"

'Subclassing / IOLEInPlaceActivate
Private WithEvents sc1 As SubClass
Attribute sc1.VB_VarHelpID = -1
Private WithEvents sc2 As SubClass
Attribute sc2.VB_VarHelpID = -1
Private m_IPAOHookStruct As IPAOHookStruct
Private IsSubClassed As Boolean

'Events
Public Event Click()
Public Event Change()
Public Event DblClick()
Public Event TabToNextControl(Reverse As Boolean)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPressByCode(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DrawItem(Index As Long, hdc As Long, bSelected As Boolean, bEnabled As Boolean, LeftPixels As Long, TopPixels As Long, RightPixels As Long, BottomPixels As Long, hFntOld As Long)

Public hListBox As Long
Private hParent As Long
Private lR As Long

Private Const LISTITEMHEIGHT = 18

Private NormalBGBrush As Long
Private SelectedBGBrush As Long
Private SelectedBorderPen As Long
Private FocusedBGBrush As Long
Private FocusedBorderPen As Long
Private SeparatorBGBrush As Long
Private SeparatorBorderPen As Long
Private GridLinesPen As Long
Private NormalTextColor As Long
Private ChosenTextColor As Long
Private ReturnCompleteTextColor As Long
Private FontNormal As Long
Private FontBold As Long
Private FontItalic As Long
Private FontStrike As Long
Private FontBoldStrike As Long
Private FontItalicStrike As Long

Public Enum eDisplayMode
    mMailingList
    mPhoneDetails
End Enum
Private mDisplayMode As eDisplayMode
Private mMultiSel As Boolean

Private Type COLUMNDIM
    cx As Long
    cWidth As Long
    cAlign As Long
End Type
Private phonedetails_columns() As COLUMNDIM
Private Const phonedetails_columnmargin = 3






'EHT=Standard
Private Sub UserControl_Initialize()
On Error GoTo ERR_HANDLER

If RunningFromIDE Then DEBUGMODE = True '[Mark]

If Not DEBUGMODE Then
    Dim a&

    hParent = UserControl.hwnd

    'Initialize phonedetails_columns
    ReDim phonedetails_columns(7)
    phonedetails_columns(0).cWidth = 25:  phonedetails_columns(0).cAlign = TA_RIGHT
    phonedetails_columns(1).cWidth = 26:  phonedetails_columns(1).cAlign = TA_CENTER
    phonedetails_columns(2).cWidth = 26:  phonedetails_columns(2).cAlign = TA_CENTER
    phonedetails_columns(3).cWidth = 275: phonedetails_columns(3).cAlign = TA_LEFT
    phonedetails_columns(4).cWidth = 97:  phonedetails_columns(4).cAlign = TA_RIGHT
    phonedetails_columns(5).cWidth = 340: phonedetails_columns(5).cAlign = TA_LEFT
    phonedetails_columns(6).cWidth = 75:  phonedetails_columns(6).cAlign = TA_LEFT
    phonedetails_columns(7).cWidth = 67:  phonedetails_columns(7).cAlign = TA_LEFT
    For a = 1 To UBound(phonedetails_columns)
        phonedetails_columns(a).cx = phonedetails_columns(a - 1).cx + phonedetails_columns(a - 1).cWidth
    Next a

    'Brushes, Pens, and Fonts (must be deleted)
    NormalBGBrush = GetSysColorBrush(COLOR_WINDOW)
    SelectedBGBrush = CreateSolidBrush(&HEBEBEB)
    SelectedBorderPen = CreatePen(PS_SOLID, 1, &H666666)
    FocusedBGBrush = CreateSolidBrush(&H40FFFF)
    FocusedBorderPen = CreatePen(PS_SOLID, 1, &H8080&)
    SeparatorBGBrush = CreateSolidBrush(&HFFDCB7)
    SeparatorBorderPen = CreatePen(PS_SOLID, 1, vbBlack)
    GridLinesPen = CreatePen(PS_SOLID, 1, &HC0C0C0)
    FontNormal = CreateFont2(UserControl.hdc, "Arial", 10, False, False, False, False)
    FontBold = CreateFont2(UserControl.hdc, "Arial", 10, True, False, False, False)
    FontItalic = CreateFont2(UserControl.hdc, "Arial", 10, False, True, False, False)
    FontStrike = CreateFont2(UserControl.hdc, "Arial", 10, False, False, False, True)
    FontBoldStrike = CreateFont2(UserControl.hdc, "Arial", 10, True, False, False, True)
    FontItalicStrike = CreateFont2(UserControl.hdc, "Arial", 10, False, True, False, True)

    'Colors
    NormalTextColor = GetSysColor(COLOR_WINDOWTEXT)
    ChosenTextColor = vbRed
    ReturnCompleteTextColor = RGB(170, 170, 170)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UserControl_Initialize", Err
End Sub

'EHT=Standard
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo ERR_HANDLER

mDisplayMode = PropBag.ReadProperty("DisplayMode", mPhoneDetails)
mMultiSel = PropBag.ReadProperty("MultiSel", False)

If Not DEBUGMODE Then
    ' Create the window:
    Dim wStyle&, lw&, lH&
    lw = UserControl.Width \ Screen.TwipsPerPixelX
    lH = UserControl.Height \ Screen.TwipsPerPixelY
    wStyle = WS_VISIBLE Or WS_CHILD Or WS_VSCROLL Or LBS_HASSTRINGS Or LBS_OWNERDRAWFIXED Or LBS_NOTIFY Or WS_HSCROLL Or LBS_SORT Or LBS_NOINTEGRALHEIGHT
    If mMultiSel Then wStyle = wStyle Or LBS_EXTENDEDSEL
    hListBox = CreateWindowEx(WS_EX_CLIENTEDGE, "listbox", "", wStyle, 0, 0, lw, lH, hParent, 0, App.hInstance, ByVal 0)
    If hListBox <> 0 Then
        ' If we succeed
        ShowWindow hListBox, SW_SHOW
        SendMessage hListBox, LB_SETITEMHEIGHT, 0, LISTITEMHEIGHT
    End If

    SetSubclassHooks
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UserControl_ReadProperties", Err
End Sub

'EHT=ResumeNext
Private Sub UserControl_Resize()
On Error Resume Next

If Not DEBUGMODE Then
    Dim lWidth As Long, lHeight As Long
    lWidth = UserControl.ScaleWidth
    lHeight = UserControl.ScaleHeight
    MoveWindow hListBox, 0, 0, lWidth, lHeight, 1
End If
End Sub

'EHT=Standard
Private Sub UserControl_Terminate()
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    UnsetSubclassHooks

    DestroyWindow hListBox
    hListBox = 0

    DeleteObject NormalBGBrush
    DeleteObject SelectedBGBrush
    DeleteObject FocusedBGBrush
    DeleteObject SelectedBorderPen
    DeleteObject FocusedBorderPen
    DeleteObject SeparatorBGBrush
    DeleteObject SeparatorBorderPen
    DeleteObject GridLinesPen
    DeleteObject FontNormal
    DeleteObject FontBold
    DeleteObject FontItalic
    DeleteObject FontStrike
    DeleteObject FontBoldStrike
    DeleteObject FontItalicStrike
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UserControl_Terminate", Err
End Sub

'EHT=Standard
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo ERR_HANDLER

PropBag.WriteProperty "DisplayMode", mDisplayMode
PropBag.WriteProperty "MultiSel", mMultiSel

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UserControl_WriteProperties", Err
End Sub


'EHT=None
Property Get DisplayMode() As eDisplayMode
DisplayMode = mDisplayMode
End Property
'EHT=None
Property Let DisplayMode(m As eDisplayMode)
mDisplayMode = m
End Property

'EHT=Standard
Property Get ItemText(i&) As String
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = SendMessage(hListBox, LB_GETTEXTLEN, i, 0)
    If lR = LB_ERR Then Exit Sub
    ItemText = Space$(lR + 1)
    lR = SendMessageByString(hListBox, LB_GETTEXT, i, ItemText)
    If lR = LB_ERR Then Exit Sub
    ItemText = Left$(ItemText, lR)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ItemText[Get]", Err
End Property

'EHT=Standard
Property Get ListCount() As Long
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    ListCount = SendMessage(hListBox, LB_GETCOUNT, 0, 0)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ListCount[Get]", Err
End Property

'EHT=Standard
Property Get ListIndex() As Long
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    ListIndex = SendMessage(hListBox, LB_GETCURSEL, 0, 0)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ListIndex[Get]", Err
End Property
'EHT=Standard
Property Let ListIndex(i&)
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    If mMultiSel Then
        lR = SendMessage(hListBox, LB_SETSEL, 0, -1)
        lR = SendMessage(hListBox, LB_SETSEL, 1, i)
    Else
        lR = SendMessage(hListBox, LB_SETCURSEL, i, 0)
        If lR <> LB_ERR Then RaiseEvent Click
    End If
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ListIndex[Let]", Err
End Property

'EHT=None
Property Get MultiSel() As Boolean
MultiSel = mMultiSel
End Property
'EHT=None
Property Let MultiSel(m As Boolean)
mMultiSel = m
End Property

'EHT=Standard
Property Get Selected(i&) As Boolean
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    Selected = SendMessage(hListBox, LB_GETSEL, i, 0)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Selected[Get]", Err
End Property
'EHT=Standard
Property Let Selected(i&, s As Boolean)
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = SendMessage(hListBox, LB_SETSEL, (Not s) + 1, i)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Selected[Let]", Err
End Property

'EHT=Standard
Property Get TopIndex() As Long
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    TopIndex = SendMessage(hListBox, LB_GETTOPINDEX, 0, 0)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "TopIndex[Get]", Err
End Property
'EHT=Standard
Property Let TopIndex(i&)
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = SendMessage(hListBox, LB_SETTOPINDEX, i, 0)
End If

Exit Property
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "TopIndex[Let]", Err
End Property






'EHT=Standard
Public Function AddItem(sectionnum&, cindex&, Optional septext$) As Long
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    'This text chosen here does not display on the screen (owner-draw), but it IS used
    '   for sorting purposes later
    If cindex < 0 Then
        If septext$ = "" Then septext$ = "-"
        AddItem = SendMessageByString(hListBox, LB_ADDSTRING, 0, sectionnum & " " & septext$)
        If AddItem <> LB_ERR Then
            SendMessage hListBox, LB_SETITEMDATA, AddItem, cindex
        End If
    Else
        With ActiveDBInstance.Clients(cindex)
            .Temp_RegenerateTempData = True
            AddItem = SendMessageByString(hListBox, LB_ADDSTRING, 0, sectionnum & " " & FormatClientName(fCustomListboxSorting, .c))
            If AddItem <> LB_ERR Then
                SendMessage hListBox, LB_SETITEMDATA, AddItem, .c.ID
            End If
        End With
    End If
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "AddItem", Err
End Function

'EHT=Standard
Public Sub Clear()
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = SendMessage(hListBox, LB_RESETCONTENT, 0, 0)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Clear", Err
End Sub

'EHT=Standard
Public Function ItemClientID(i&) As Long
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    ItemClientID = SendMessage(hListBox, LB_GETITEMDATA, i, 0)
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "ItemClientID", Err
End Function

'EHT=Standard
Public Sub RemoveItem(i&)
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = SendMessage(hListBox, LB_DELETESTRING, i&, 0)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "RemoveItem", Err
End Sub

'EHT=Standard
Public Sub Repaint()
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = InvalidateRect(hListBox, 0, 1)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Repaint", Err
End Sub

'EHT=Standard
Public Function SelectedClientID() As Long
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    SelectedClientID = ListIndex
    If SelectedClientID <> LB_ERR Then SelectedClientID = ItemClientID(SelectedClientID)
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SelectedClientID", Err
End Function

'EHT=Standard
Public Sub SetRedraw(r As Boolean)
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    lR = SendMessage(hListBox, WM_SETREDRAW, (Not r) + 1, 0)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetRedraw", Err
End Sub


'EHT=Standard
Private Sub DrawItem_MailingList(cindex&, dis As DRAWITEMSTRUCT)
On Error GoTo ERR_HANDLER

Dim tc&, nx&, ny&, t$

If ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData Then tabSearch.RegenerateClientTempData cindex

With ActiveDBInstance.Clients(cindex).c
    If frmMain.CHOS_IsChosen(dis.ItemData) Then
        tc = ChosenTextColor
    ElseIf Flag_IsSet(.LastYear_Flags, CompletedReturn) Then
        tc = ReturnCompleteTextColor
    Else
        tc = NormalTextColor
    End If
    SetTextColor dis.hdc, tc
    SelectObject dis.hdc, FontNormal
    SetTextAlign dis.hdc, TA_LEFT

    nx = dis.rcItem.Left + 3
    ny = dis.rcItem.Top + 1

    t$ = FormatClientName(fMailingList, ActiveDBInstance.Clients(cindex).c)
    TextOut dis.hdc, nx, ny, t$, Len(t$)
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DrawItem_MailingList", Err
End Sub

'EHT=Standard
Private Sub DrawItem_PhoneDetails(cindex&, dis As DRAWITEMSTRUCT)
On Error GoTo ERR_HANDLER

Dim tc&, p$(), a&, ci&, c As Boolean, nx&, ny&, t$, atleastoneliving As Boolean

If ActiveDBInstance.Clients(cindex).Temp_RegenerateTempData Then tabSearch.RegenerateClientTempData cindex

With ActiveDBInstance.Clients(cindex).c
    atleastoneliving = ((Len(.Person1.First) > 0) And (.Person1.dod = NullLong)) Or ((Len(.Person2.First) > 0) And (.Person2.dod = NullLong))
    If frmMain.CHOS_IsChosen(dis.ItemData) Then
        tc = ChosenTextColor
    ElseIf Flag_IsSet(.Flags, CompletedReturn) Then
        tc = ReturnCompleteTextColor
    ElseIf Flag_IsSet(.Flags, NoNeedToFile) Then
        tc = ReturnCompleteTextColor
    ElseIf atleastoneliving Then
        'At least one person still living
        tc = NormalTextColor
    Else
        'All deceased
        tc = ReturnCompleteTextColor
    End If
    SetTextColor dis.hdc, tc
    SelectObject dis.hdc, FontNormal

    nx = dis.rcItem.Left + 3
    ny = dis.rcItem.Top + 1

    'Draw grid lines
    SelectObject dis.hdc, GridLinesPen
    For a = 1 To UBound(phonedetails_columns)
        MoveToEx dis.hdc, nx + phonedetails_columns(a).cx, dis.rcItem.Top, 0
        LineTo dis.hdc, nx + phonedetails_columns(a).cx, dis.rcItem.Bottom
    Next a

    'LYMinutes
    ci = 0
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    If Flag_IsSet(.Flags, NewClient) Then
        t$ = ""
    ElseIf Flag_IsSet(.LastYear_Flags, NoNeedToFile) Then
        t$ = "NF"
    Else
        t$ = FieldToString(.LastYear_MinutesToComplete, mNumberOrNULL)
    End If
    TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)

    'NumApptSlots
    ci = ci + 1
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    t$ = FormatNumApptSlots(.NumApptSlotsToUse)
    TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)

    'Last year's DO/MI flag
    ci = ci + 1
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    If Flag_IsSet(.LastYear_Flags, DroppedOff) Then
        t$ = "DO"
        TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)
    ElseIf Flag_IsSet(.LastYear_Flags, MailedIn) Then
        t$ = "MI"
        TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)
    End If

    'Last, First
    ci = ci + 1
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign Or TA_UPDATECP
    MoveToEx dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, 0
    p$ = Split(ActiveDBInstance.Clients(cindex).Temp_ParsedName, BoldSep)
    For a = 0 To UBound(p$)
        If c Then
            If atleastoneliving Then
                SelectObject dis.hdc, FontBold
            Else
                SelectObject dis.hdc, FontBoldStrike
            End If
            c = False
        Else
            If atleastoneliving Then
                SelectObject dis.hdc, FontNormal
            Else
                SelectObject dis.hdc, FontStrike
            End If
            c = True
        End If
        TextOut dis.hdc, 0, 0, p$(a), Len(p$(a))
    Next a
    SelectObject dis.hdc, FontNormal
    SetTextAlign dis.hdc, TA_NOUPDATECP

    'Phone numbers
    ci = ci + 1
    t$ = FieldToString(.PhoneHome, mPhoneHideLocalAreaCode)
    a = InStr(t$, "x")
    If a > 0 Then
        SetTextAlign dis.hdc, TA_LEFT
        TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)) - 5, ny, "x", 1

        t$ = Left$(t$, a - 1)
    End If
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)) - 6, ny, t$, Len(t$)

    'Notes
    ci = ci + 1
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    SelectObject dis.hdc, FontBold
    t$ = .Notes
    TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)
    SelectObject dis.hdc, FontNormal

    'Appointment
    ci = ci + 1
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    If ActiveDBInstance.Clients(cindex).Temp_ApptPast Then
        If ActiveDBInstance.Clients(cindex).Temp_DidntHappen Then
            SelectObject dis.hdc, FontItalicStrike
        Else
            SelectObject dis.hdc, FontItalic
        End If
    Else
        If ActiveDBInstance.Clients(cindex).Temp_DidntHappen Then
            SelectObject dis.hdc, FontStrike
        Else
            SelectObject dis.hdc, FontNormal
        End If
    End If
    t$ = ActiveDBInstance.Clients(cindex).Temp_ApptDate
    TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)
    SelectObject dis.hdc, FontNormal

    'Current flags
    ci = ci + 1
    SetTextAlign dis.hdc, phonedetails_columns(ci).cAlign
    '(Pos 1)
    If Flag_IsSet(.Flags, NewClient) Then
        t$ = "NN"
        TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)), ny, t$, Len(t$)
    End If
    '(Pos 2)
    t$ = ""
    If Flag_IsSet(.Flags, DroppedOff) Then
        t$ = "DO"
    ElseIf Flag_IsSet(.Flags, MailedIn) Then
        t$ = "MI"
    ElseIf Flag_IsSet(.Flags, PartiallyComplete) Then
        t$ = "Inc"
    End If
    If t$ <> "" Then TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)) + 22, ny, t$, Len(t$)
    '(Pos 3)
    t$ = ""
    If Flag_IsSet(.Flags, ReleasedBeforePayment) Then
        t$ = "Rel"
    ElseIf Flag_IsSet(.Flags, NoNeedToFile) Then
        t$ = "NF"
    End If
    If t$ <> "" Then TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)) + 44, ny, t$, Len(t$)
    '(Pos 4)
    t$ = ""
    If Flag_IsSet(.Flags, Extension) Then
        t$ = "Ext"
        TextOut dis.hdc, nx + GetTextDrawPos(phonedetails_columns(ci)) + 66, ny, t$, Len(t$)
    End If
End With

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DrawItem_PhoneDetails", Err
End Sub

'EHT=Standard
Private Sub DrawItem_Separator(sepid&, dis As DRAWITEMSTRUCT)
On Error GoTo ERR_HANDLER

'The sepid parametor isn't used yet, but leave it in case it's useful in the future

Dim nx&, ny&, t$

SetTextColor dis.hdc, NormalTextColor
SelectObject dis.hdc, FontBold
SetTextAlign dis.hdc, TA_CENTER

'nx = dis.rcItem.Left + 3
nx = dis.rcItem.Left + ((dis.rcItem.Right - dis.rcItem.Left) / 2)
ny = dis.rcItem.Top + 1

t$ = Mid$(ItemText(dis.ItemId), 3)
TextOut dis.hdc, nx, ny, t$, Len(t$)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DrawItem_Separator", Err
End Sub















'EHT=Standard
Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ERR_HANDLER

Dim dis As DRAWITEMSTRUCT, cindex&
CopyMemory dis, ByVal lParam, Len(dis)
With dis
    If .ItemId >= 0 Then
        Select Case .ItemAction
        Case ODA_SELECT, ODA_FOCUS, ODA_DRAWENTIRE
            SetBkMode .hdc, TRANSPARENT

            If .ItemData < 0 Then
                'Separator item
                SelectObject .hdc, SeparatorBGBrush
                SelectObject .hdc, SeparatorBorderPen
                Rectangle .hdc, .rcItem.Left - 1, .rcItem.Top, .rcItem.Right + 1, .rcItem.Bottom
                DrawItem_Separator .ItemData, dis
            Else
                'Client item
                If Flag_IsSet(.ItemState, ODS_SELECTED) Then
                    'Draw selection rectangle
                    If Flag_IsSet(.ItemState, ODS_FOCUS) Then
                        SelectObject .hdc, FocusedBGBrush
                        SelectObject .hdc, FocusedBorderPen
                    Else
                        SelectObject .hdc, SelectedBGBrush
                        SelectObject .hdc, SelectedBorderPen
                    End If
                    Rectangle .hdc, .rcItem.Left - 1, .rcItem.Top, .rcItem.Right + 1, .rcItem.Bottom
                Else
                    'Clear background
                    FillRect .hdc, .rcItem, NormalBGBrush
                End If

                'Draw item
                cindex = DB_FindClientIndex(ActiveDBInstance, .ItemData)
                If cindex >= 0 Then
                    Select Case mDisplayMode
                    Case mMailingList:  DrawItem_MailingList cindex, dis
                    Case mPhoneDetails: DrawItem_PhoneDetails cindex, dis
                    End Select
                End If
            End If
        End Select
    End If
End With

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "DrawItem", Err
End Function
'EHT=Standard
Private Function GetTextDrawPos(c As COLUMNDIM) As Long
On Error GoTo ERR_HANDLER

Select Case c.cAlign
Case TA_LEFT
    GetTextDrawPos = c.cx + phonedetails_columnmargin
Case TA_CENTER
    GetTextDrawPos = c.cx + (c.cWidth / 2)
Case TA_RIGHT
    GetTextDrawPos = c.cx + c.cWidth + 1 - phonedetails_columnmargin
End Select

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "GetTextDrawPos", Err
End Function

'EHT=Standard
Private Function sc1_WindowProc(hwnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
On Error GoTo ERR_HANDLER

Select Case uMsg
Case WM_COMMAND
    If (wParam \ &H10000) = LBN_DBLCLK Then RaiseEvent DblClick
    sc1_WindowProc = 1

Case WM_DRAWITEM
    sc1_WindowProc = DrawItem(wParam, lParam)

Case WM_SETFOCUS
    SetFocusAPI hListBox
End Select

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "sc1_WindowProc", Err
End Function

'EHT=Standard
Private Function sc2_WindowProc(hwnd As Long, uMsg As Long, wParam As Long, lParam As Long) As Long
On Error GoTo ERR_HANDLER

Dim iKeyCode As Integer, iShift As Integer
Dim iButton As Integer, X As Single, Y As Single
Select Case uMsg
Case WM_KEYDOWN
    iKeyCode = (wParam And &HFF)
    If iKeyCode <> 0 Then
        iShift = GetShiftState
        If iKeyCode = vbKeyTab And (iShift = 0 Or iShift = vbShiftMask) Then
            RaiseEvent TabToNextControl(iShift = vbShiftMask)
        Else
            RaiseEvent KeyDown(iKeyCode, iShift)
            If iKeyCode <> 0 Then   'iKeyCode=0 means the event was cancelled, so don't pass it on
                wParam = (wParam And Not &HFF&) Or (iKeyCode And &HFF&)
                sc2.CallOldWndProc
            End If
        End If
    End If
Case WM_CHAR
    iKeyCode = (wParam And &HFF)
    If iKeyCode <> 0 Then
        If iKeyCode = vbKeyTab And (iShift = 0 Or iShift = vbShiftMask) Then
            'Eat the tab keystrokes
        Else
            RaiseEvent KeyPressByCode(iKeyCode, GetShiftState())
            If iKeyCode <> 0 Then   'iKeyCode=0 means the event was cancelled, so don't pass it on
                wParam = (wParam And Not &HFF&) Or (iKeyCode And &HFF&)
                sc2.CallOldWndProc
            End If
        End If
    End If
Case WM_KEYUP
    iKeyCode = (wParam And &HFF)
    If iKeyCode <> 0 Then
        If iKeyCode = vbKeyTab And (iShift = 0 Or iShift = vbShiftMask) Then
            'Eat the tab keystrokes
        Else
            RaiseEvent KeyUp(iKeyCode, GetShiftState())
            If iKeyCode <> 0 Then   'iKeyCode=0 means the event was cancelled, so don't pass it on
                wParam = (wParam And Not &HFF&) Or (iKeyCode And &HFF&)
                sc2.CallOldWndProc
            End If
        End If
    End If

Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
    iButton = (Abs(uMsg = WM_LBUTTONDOWN)) * vbLeftButton + (Abs(uMsg = WM_RBUTTONDOWN)) * vbRightButton + (Abs(uMsg = WM_MBUTTONDOWN)) * vbMiddleButton
    iShift = wParam And &HC&  'Only the Ctrl (0x8) and Shift (0x4) flags are passed
    If iShift > 0 Then iShift = iShift / 4      'Shift it by 2 bits to conform to VB6's shift codes
    If (lParam And &H8000&) = &H8000& Then
        X = -(&H8000& - (lParam And &H7FFF&))
    Else
        X = (lParam And &HFFFF&)
    End If
    If (lParam And &H80000000) = &H80000000 Then
        Y = -(&H8000& - (lParam And &H7FFF0000) \ &H10000)
    Else
        Y = (lParam \ &H10000)
    End If
    RaiseEvent MouseDown(iButton, iShift, X, Y)
Case WM_MOUSEMOVE
    iButton = Abs(GetAsyncKeyState(vbKeyLButton) <> 0) * vbLeftButton + Abs(GetAsyncKeyState(vbKeyRButton) <> 0) * vbRightButton + Abs(GetAsyncKeyState(vbKeyMButton) <> 0) * vbMiddleButton
    iShift = wParam And &HC&    'Only the Ctrl (0x8) and Shift (0x4) flags are passed
    If iShift > 0 Then iShift = iShift / 4      'Shift it by 2 bits to conform to VB6's shift codes
    If (lParam And &H8000&) = &H8000& Then
        X = -(&H8000& - (lParam And &H7FFF&))
    Else
        X = (lParam And &HFFFF&)
    End If
    If (lParam And &H80000000) = &H80000000 Then
        Y = -(&H8000& - (lParam And &H7FFF0000) \ &H10000)
    Else
        Y = (lParam \ &H10000)
    End If
    RaiseEvent MouseMove(iButton, iShift, X, Y)
Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
    iButton = (Abs(uMsg = WM_LBUTTONUP)) * vbLeftButton + (Abs(uMsg = WM_RBUTTONUP)) * vbRightButton + (Abs(uMsg = WM_MBUTTONUP)) * vbMiddleButton
    iShift = wParam And &HC&    'Only the Ctrl (0x8) and Shift (0x4) flags are passed
    If iShift > 0 Then iShift = iShift / 4      'Shift it by 2 bits to conform to VB6's shift codes
    If (lParam And &H8000&) = &H8000& Then
        X = -(&H8000& - (lParam And &H7FFF&))
    Else
        X = (lParam And &HFFFF&)
    End If
    If (lParam And &H80000000) = &H80000000 Then
        Y = -(&H8000& - (lParam And &H7FFF0000) \ &H10000)
    Else
        Y = (lParam \ &H10000)
    End If
    RaiseEvent MouseUp(iButton, iShift, X, Y)

Case WM_MOUSEACTIVATE
    If GetFocus() <> hListBox Then
        SetFocusAPI UserControl.hwnd
        sc2_WindowProc = MA_NOACTIVATE
    Else
        sc2.CallOldWndProc
    End If

Case WM_SETFOCUS
    'Get in-place frame and make sure it is set to our in-between
    'implementation of IOleInPlaceActiveObject in order to catch
    'TranslateAccelerator calls
    Dim pOleObject                  As IOleObject
    Dim pOleInPlaceSite             As IOleInPlaceSite
    Dim pOleInPlaceFrame            As IOleInPlaceFrame
    Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
    Dim pOleInPlaceActiveObject     As IOleInPlaceActiveObject
    Dim PosRect                     As RECT
    Dim ClipRect                    As RECT
    Dim FrameInfo                   As OLEINPLACEFRAMEINFO
    Set pOleObject = Me
    Set pOleInPlaceSite = pOleObject.GetClientSite
    pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
    CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
    pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
    If Not pOleInPlaceUIWindow Is Nothing Then
        pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
    End If
    CopyMemory pOleInPlaceActiveObject, 0&, 4
End Select

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "sc2_WindowProc", Err
End Function

'EHT=Standard
Private Sub SetSubclassHooks()
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    Set sc1 = New SubClass
    sc1.HookMessage WM_COMMAND, cBefore
    sc1.HookMessage WM_DRAWITEM, cBefore
    sc1.HookMessage WM_SETFOCUS, cBefore
    sc1.SetHook hParent

    Set sc2 = New SubClass
    sc2.HookMessage WM_KEYDOWN, cManual
    sc2.HookMessage WM_CHAR, cManual
    sc2.HookMessage WM_KEYUP, cManual
    sc2.HookMessage WM_LBUTTONDOWN, cBefore
    sc2.HookMessage WM_MBUTTONDOWN, cBefore
    sc2.HookMessage WM_RBUTTONDOWN, cBefore
    sc2.HookMessage WM_MOUSEMOVE, cBefore
    sc2.HookMessage WM_LBUTTONUP, cBefore
    sc2.HookMessage WM_MBUTTONUP, cBefore
    sc2.HookMessage WM_RBUTTONUP, cBefore
    sc2.HookMessage WM_MOUSEACTIVATE, cManual
    sc2.HookMessage WM_SETFOCUS, cBefore
    sc2.SetHook hListBox
    IsSubClassed = True
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetSubclassHooks", Err
End Sub
'EHT=Standard
Private Sub UnsetSubclassHooks()
On Error GoTo ERR_HANDLER

If Not DEBUGMODE Then
    If IsSubClassed Then
        sc1.UnSetHook
        sc2.UnSetHook
        IsSubClassed = False
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UnsetSubclassHooks", Err
End Sub

'EHT=None
Friend Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
TranslateAccelerator = S_FALSE
End Function
