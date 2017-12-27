Attribute VB_Name = "modCommon"
Option Explicit
Private Const MOD_NAME = "modCommon"

'This module should NEVER reference ActiveDBInstance (use LocalDBInstance instead)

Public Const BLACK_BRUSH = 4
Public Const BLACK_PEN = 7
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400                'Determines the width and height of the rectangle.
Public Const DT_CENTER = &H1
Public Const DT_EDITCONTROL = &H2000            'Duplicates the text-displaying characteristics of a multiline edit control.
Public Const DT_END_ELLIPSIS = &H8000           'If the end of a string does not fit in the rectangle, it is truncated and ellipses are added.
Public Const DT_EXPANDTABS = &H40               'Expands tab characters.
Public Const DT_EXTERNALLEADING = &H200         'Includes the font external leading in line height.
Public Const DT_LEFT = &H0
Public Const DT_HIDEPREFIX = &H100000           'Ignores the ampersand (&) prefix character in the text.
Public Const DT_INTERNAL = &H1000               'Uses the system font to calculate text metrics.
Public Const DT_MODIFYSTRING = &H10000          'Modifies the specified string to match the displayed text.
Public Const DT_NOCLIP = &H100                  'Draws without clipping.
Public Const DT_NOFULLWIDTHCHARBREAK = &H80000  'Prevents a line break at a DBCS (double-wide character string),
Public Const DT_NOPREFIX = &H800
Public Const DT_PATH_ELLIPSIS = &H4000          'Replaces characters in the middle of the string with ellipses
Public Const DT_PREFIXONLY = &H200000
Public Const DT_RIGHT = &H2
Public Const DT_RTLREADING = &H20000            'Layout in right-to-left reading order for bidirectional text when the font selected into the hdc is a Hebrew or Arabic font.
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80                  'Change the default number (8) of characters per tab
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const FW_BOLD = 700
Public Const FW_NORMAL = 400
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const GRAY_BRUSH = 2
Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = (-4)
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const LB_ADDFILE = &H196
Public Const LB_ADDSTRING = &H180
Public Const LB_CTLCODE = 0&
Public Const LB_DELETESTRING = &H182
Public Const LB_DIR = &H18D
Public Const LB_ERR = (-1)
Public Const LB_ERRSPACE = (-2)
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETANCHORINDEX = &H19D
Public Const LB_GETCARETINDEX = &H19F
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETCURSEL = &H188
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_GETITEMRECT = &H198
Public Const LB_GETLOCALE = &H1A6
Public Const LB_GETSEL = &H187
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETSELITEMS = &H191
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_INSERTSTRING = &H181
Public Const LB_ITEMFROMPOINT = &H1A9
Public Const LB_MSGMAX = &H1A8
Public Const LB_OKAY = 0
Public Const LB_RESETCONTENT = &H184
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_SELITEMRANGEEX = &H183
Public Const LB_SETANCHORINDEX = &H19C
Public Const LB_SETCARETINDEX = &H19E
Public Const LB_SETCOLUMNWIDTH = &H195
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETITEMDATA = &H19A
Public Const LB_SETITEMHEIGHT = &H1A0
Public Const LB_SETLOCALE = &H1A5
Public Const LB_SETSEL = &H185
Public Const LB_SETTABSTOPS = &H192
Public Const LB_SETTOPINDEX = &H197
Public Const LBN_DBLCLK = 2
Public Const LBN_ERRSPACE = (-2)
Public Const LBN_KILLFOCUS = 5
Public Const LBN_SELCANCEL = 3
Public Const LBN_SELCHANGE = 1
Public Const LBN_SETFOCUS = 4
Public Const LBS_DISABLENOSCROLL = &H1000&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_NODATA = &H2000&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_NOTIFY = &H1&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_SORT = &H2&
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_WANTKEYBOARDINPUT = &H400&
Public Const LF_FACESIZE = 32
Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3
Public Const MA_NOACTIVATEANDEAT = 4
Public Const NULL_BRUSH = 5
Public Const NULL_PEN = 8
Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_FOCUS = &H4
Public Const ODA_SELECT = &H2
Public Const ODS_CHECKED = &H8
Public Const ODS_COMBOBOXEDIT = &H1000
Public Const ODS_DEFAULT = &H20
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2
Public Const ODS_HOTLIGHT = &H40
Public Const ODS_INACTIVE = &H80
Public Const ODS_SELECTED = &H1
Public Const ODT_BUTTON = 4
Public Const ODT_COMBOBOX = 3
Public Const ODT_HEADER = 100
Public Const ODT_LISTBOX = 2
Public Const ODT_LISTVIEW = 102
Public Const ODT_MENU = 1
Public Const ODT_STATIC = 5
Public Const ODT_TAB = 101
Public Const OPAQUE = 2
Public Const PS_DASH = 1
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_DOT = 2
Public Const PS_NULL = 5
Public Const PS_SOLID = 0
Public Const SPI_GETWORKAREA = 48
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const TA_BASELINE = 24
Public Const TA_BOTTOM = 8
Public Const TA_CENTER = 6
Public Const TA_LEFT = 0
Public Const TA_NOUPDATECP = 0
Public Const TA_RIGHT = 2
Public Const TA_TOP = 0
Public Const TA_UPDATECP = 1
Public Const TRANSPARENT = 1
Public Const VK_CAPITAL = &H14
Public Const VK_CONTROL = &H11&
Public Const VK_MENU = &H12& ' Alt key
Public Const VK_SHIFT = &H10&
Public Const WH_CALLWNDPROC = 4
Public Const WHITE_BRUSH = 0
Public Const WHITE_PEN = 6
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_DRAWITEM = &H2B
Public Const WM_ERASEBKGND = &H14
Public Const WM_GETFONT = &H31
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_INITDIALOG = &H110
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MEASUREITEM = &H2
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NCPAINT = &H85
Public Const WM_PAINT = &HF
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETFOCUS = &H7
Public Const WM_SETFONT = &H30
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SIZE = &H5
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOZORDER = &H4

Public Const SEP1 = "|"
Public Const SEP2 = ","
Public Const MultiLineSep = "|"
Public Const BoldSep = "|"

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type
Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type SIZE
    cx As Long
    cy As Long
End Type
Public Type DRAWITEMSTRUCT
    Ctltype As Long
    CtlID As Long
    ItemId As Long
    ItemAction As Long
    ItemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type
Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE) As Byte
End Type
Public Type ICONLISTBOXITEMINFO
    lItemData As Long           ' Provide item data - normal item data is used to store a pointer to a structure of this public type
    lExtraData As Long          ' An additional item data
    lIconIndex As Long          ' Index of icon in icon list, if required
    lIndentSize As Long         ' How far the text should be indented from left, in pixels
    lItemHeight As Long         ' How high a single item should be
    lForeColour As OLE_COLOR    ' Fore colour of the item
    lBackColour As OLE_COLOR    ' Back colour of the item
    bUnderLineItem As Boolean   ' Whether a ruling should be placed below the item
    bOverLineItem As Boolean    ' Whether a ruling should be placed above the item
    dFontSize As Single         ' VB font size, stored here for ease of extracting a font object
    tLF As LOGFONT              ' API font description.  lfFaceName should have all bytes = 0 to use default
    lTextAlignX As Long         ' Horizonal Text alignment
    lTextAlignY As Long         ' Vertical Text alignment
End Type

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByRect Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As RECT) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetDCBrushColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetDCPenColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub CopyMemoryByLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DwmIsCompositionEnabled Lib "dwmapi" (pfEnabled As Long) As Long


Public Enum ClientNameFormatMode
    fSearchResults
    fMailingList
    fExport_First
    fExport_Last
    fExport_Display
    fPrintLabels
    fSchedulePct
    fPullFiles
    fLog
    fCustomListboxSorting
    fChosenClients
    fUnpaid
End Enum

Public Enum FieldFormatMode
    mNumber = 10
    mNumber_NoNegativeNumber = 13
    mNumberOrNULL = 11
    mNumberOrNULL_NoNegativeNumber = 12
    mDollar = 20
    mDollarOrNULL = 21
    mDollarOrNULL_NoNegativeNumber = 23
    mDollarOrNULL_ZeroForcedToNullLong = 22
    mDollarOrNULL_ZeroForcedToNullLong_NoNegativeNumber = 24
    mDateAsLong = 30
    mDateAsLongOrNULL = 31
    mTime = 40
    mYearOrNULL = 70
    mString = 50
    mStringUC = 51
    mStringLC = 52
    mStringMultiline = 53
    mStringCommaSeparatedStateList = 54
    mPhone = 60
    mPhoneHideLocalAreaCode = 61
End Enum

Public Enum ScheduleShapeStyle
    Style_Normal
    Style_New
    Style_Move
    Style_MoveAndCtrlCopy
    Style_Copy
    Style_CopyForcedWithCtrl
    Style_ShowAppt
End Enum

Public Const NullLong = &H80000000      '-2147483648
Public Const NullDouble = -1.79769313486231E+308

Private MouseNullZoneRef As POINTAPI
Private MouseNullZonePixels&

'EHT=None
Sub AddOpNote(ByRef OpNotes$, ByVal newon$)
newon$ = Format$(Now, "yyyy-mm-dd hh:mma/p") & " " & newon$
If Len(OpNotes$) = 0 Then
    OpNotes$ = newon$
Else
    OpNotes$ = OpNotes$ & vbCrLf & newon$
End If
End Sub

'EHT=None
Function CapatalizeFirstLetter(s$) As String
CapatalizeFirstLetter = UCase$(Left$(s$, 1)) & LCase$(Mid$(s$, 2))
End Function

'EHT=None
Function CreateFont2(fHDC&, fName$, fSize!, fBold As Boolean, fItalic As Boolean, fUnderline As Boolean, fStrikeout As Boolean) As Long
Dim fw&
If fBold Then fw = FW_BOLD Else fw = FW_NORMAL
CreateFont2 = CreateFont(-(fSize * GetDeviceCaps(fHDC, LOGPIXELSY)) / 72, 0, 0, 0, fw, fItalic, fUnderline, fStrikeout, 0, 0, 0, 0, 0, fName$)
End Function

'EHT=None
Sub EnableTextbox(txt As TextBox, e As Boolean)
txt.Enabled = e
If e Then
    txt.BackColor = vbWindowBackground
Else
    txt.BackColor = vbButtonFace
End If
End Sub

'EHT=Standard
Function FieldFromString(s$, m As FieldFormatMode) As Variant
On Error GoTo ERR_HANDLER

Dim v$, a&, n$, c$, nv#, e$, cy&, d$(), dv&(2)
v$ = Trim$(s$)

Select Case m
'Long
Case mNumber, mNumber_NoNegativeNumber, _
     mDollar
    FieldFromString = Val(Replace$(Replace$(v$, "$", ""), ",", ""))
    If m = mNumber_NoNegativeNumber Then
        If FieldFromString < 0 Then FieldFromString = 0
    End If
Case mNumberOrNULL, mNumberOrNULL_NoNegativeNumber, _
     mDollarOrNULL, mDollarOrNULL_NoNegativeNumber, _
     mDollarOrNULL_ZeroForcedToNullLong, mDollarOrNULL_ZeroForcedToNullLong_NoNegativeNumber
    If Len(v$) = 0 Then
        FieldFromString = NullLong
    Else
        FieldFromString = Val(Replace$(Replace$(v$, "$", ""), ",", ""))
        If (m = mNumberOrNULL_NoNegativeNumber) Or (m = mDollarOrNULL_NoNegativeNumber) Or (m = mDollarOrNULL_ZeroForcedToNullLong_NoNegativeNumber) Then
            If FieldFromString < 0 Then FieldFromString = 0
        End If
        If (m = mDollarOrNULL_ZeroForcedToNullLong) Or (m = mDollarOrNULL_ZeroForcedToNullLong_NoNegativeNumber) Then
            If FieldFromString = 0 Then FieldFromString = NullLong
        End If
    End If

'Date stored in a Long
Case mDateAsLong, mDateAsLongOrNULL
    cy = Year(Date)
    If IsNumeric(v$) Then
        Select Case Len(v$)
        Case 4, 6, 8:                               'MMDD or MMDDYY or MMDDYYYY
            dv(0) = Mid$(v$, 1, 2)                  'Month
            dv(1) = Mid$(v$, 3, 2)                  'Day
            If Len(v$) > 4 Then
                dv(2) = Mid$(v$, 5, Len(v$) - 4)    'Year
            Else
                dv(2) = cy                          'Assume current year
            End If
        End Select
    ElseIf IsNumeric(Replace$(v$, "/", "")) Then
        d$ = Split(v$, "/")
        If UBound(d$) >= 1 Then                     'M/D or M/D/Y
            dv(0) = d$(0)
            dv(1) = d$(1)
            If UBound(d$) >= 2 Then
                dv(2) = d$(2)                       'M/D/Y
            Else
                dv(2) = cy                          'M/D, assume current year
            End If
        End If
    End If
    If dv(0) > 0 Then
        If dv(2) < 100 Then
            If dv(2) > (cy Mod 100) Then
                'If user typed a 2-digit year that is after the current year, assume he meant the 1900s
                dv(2) = ((cy \ 100) * 100) - 100 + dv(2)
            Else
                'Otherwise put in the current century
                dv(2) = ((cy \ 100) * 100) + dv(2)
            End If
        End If
        FieldFromString = DateSerial(dv(2), dv(0), dv(1))
    ElseIf m = mDateAsLong Then
        FieldFromString = 0
    Else
        FieldFromString = NullLong
    End If

'Date
Case mTime
    If IsDate(v$) Then
        FieldFromString = TimeValue(v$)
    Else
        FieldFromString = 0
    End If

'Long
Case mYearOrNULL
    If Len(v$) = 0 Then
        FieldFromString = NullLong
    Else
        FieldFromString = Val(Replace$(Replace$(v$, "$", ""), ",", ""))
    End If

'String
Case mString
    FieldFromString = FormatTextForDB(v$)
Case mStringUC
    FieldFromString = UCase$(v$)
Case mStringLC
    FieldFromString = LCase$(v$)
Case mStringMultiline
    FieldFromString = v$
Case mStringCommaSeparatedStateList
    v$ = UCase$(v$)
    For a = 1 To Len(v$)
        c$ = Mid$(v$, a, 1)
        If Asc(c$) >= 65 And Asc(c$) <= 90 Then
            n$ = n$ & c$
        End If
    Next a
    v$ = ""
    For a = 1 To Len(n$) Step 2
        v$ = v$ & Mid$(n$, a, 2) & ","
    Next a
    If Len(v$) > 0 Then
        FieldFromString = Left$(v$, Len(v$) - 1)
    Else
        FieldFromString = v$
    End If

'String
Case mPhone
    For a = 1 To Len(v$)
        c$ = Mid$(v$, a, 1)
        If LCase$(c$) = "x" Then
            Exit For
        ElseIf IsNumeric(c$) Then
            n$ = n$ & c$
            If Len(n$) >= 10 Then Exit For     'Allow only up to 10 digits
        End If
    Next a
    a = InStr(a, LCase$(v$), "x")
    If (n$ = "") And (a = 0) Then
        FieldFromString = ""
    Else
        If a > 0 Then e$ = "x" & UCase$(Mid$(v$, a + 1))
        If Len(n$) = 7 Then n$ = DB_GetSetting(ActiveDBInstance, "GLOBAL_LocalAreaCode") & n$
        nv = Val(n$)
        FieldFromString = Format$(nv, "0000000000") & e$
    End If

Case Else
    FieldFromString = v$
    Err.Raise 1, , "Unknown FieldFormatMode"
End Select

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "FieldFromString", Err
End Function

'EHT=None
Sub FieldFromTextbox(txt As TextBox, ByRef v As Variant)
'Converts display format to database format
v = FieldFromString(txt.Text, Val(txt.Tag))
End Sub

'EHT=Standard
Function FieldToString(v As Variant, m As FieldFormatMode) As String
On Error GoTo ERR_HANDLER

Dim t$, a&

Select Case m
'Long
Case mNumber, mNumber_NoNegativeNumber
    FieldToString = Format$(v, "#,##0")
Case mNumberOrNULL, mNumberOrNULL_NoNegativeNumber
    If v <> NullLong Then FieldToString = Format$(v, "#,##0")

'Long
Case mDollar
    FieldToString = Format$(v, "$#,##0")
Case mDollarOrNULL, mDollarOrNULL_NoNegativeNumber, mDollarOrNULL_ZeroForcedToNullLong, mDollarOrNULL_ZeroForcedToNullLong_NoNegativeNumber
    If v <> NullLong Then FieldToString = Format$(v, "$#,##0")

'Date stored in a Long
Case mDateAsLong
    FieldToString = Format$(v, "m/dd/yyyy")
Case mDateAsLongOrNULL
    If v <> NullLong Then FieldToString = Format$(v, "m/dd/yyyy")

'Date
Case mTime
    FieldToString = Format$(v, "h:mm AM/PM")

'Long
Case mYearOrNULL
    If v <> NullLong Then FieldToString = Format$(v, "0000")

'String
Case mString, mStringMultiline, mStringCommaSeparatedStateList
    FieldToString = v
Case mStringUC
    FieldToString = UCase$(v)
Case mStringLC
    FieldToString = LCase$(v)

'String
Case mPhone, mPhoneHideLocalAreaCode
    t$ = CStr(v)
    If t$ <> "" Then
        a = InStr(LCase$(t$), "x")
        If a = 0 Then a = Len(t$) + 1
        If a = 11 Then
            If m = mPhoneHideLocalAreaCode Then
                If Mid$(t$, 1, 3) <> DB_GetSetting(ActiveDBInstance, "GLOBAL_LocalAreaCode") Then
                    FieldToString = "(" & Mid$(t$, 1, 3) & ") "
                End If
            Else
                FieldToString = "(" & Mid$(t$, 1, 3) & ") "
            End If
            FieldToString = FieldToString & Mid$(t$, 4, 3) & "-" & Mid$(t$, 7)
        Else
            FieldToString = Mid$(t$, 1, 3) & "-" & Mid$(t$, 4)
        End If
    End If

Case Else
    FieldToString = v
    Err.Raise 1, , "Unknown FieldFormatMode"
End Select

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "FieldToString", Err
End Function

'EHT=None
Sub FieldToTextbox(txt As TextBox, v As Variant, Optional en As Boolean = -100)
'Converts database format to display format
txt.Text = FieldToString(v, Val(txt.Tag))
If en <> -100 Then EnableTextbox txt, en
End Sub

'EHT=None
Function Flag_IsSet(Flags As Long, Flag As Long) As Boolean
Flag_IsSet = ((Flags And Flag) = Flag)
End Function

'EHT=None
Function Flag_Remove(Flags As Long, Flag As Long) As Long
Flag_Remove = Not ((Not Flags) Or Flag)
End Function

'EHT=None
Function Flag_ToCheckbox(Flags As Long, Flag As Long) As Integer
Dim b As Boolean
b = ((Flags And Flag) = Flag)     'Is flag set?
Flag_ToCheckbox = (Not b) + 1    'TRUE = 1, FALSE = 0
End Function

'EHT=None
Function FormatApptTime$(Day&, actualtime As Date)
FormatApptTime$ = Format$(CDate(Day) + actualtime, "m/dd h:mma/p")
End Function

'EHT=None
Function FormatApptTime2$(at As Date)
FormatApptTime2$ = Format$(at, "m/dd h:mma/p")
End Function

'EHT=None
Function FormatDateForDayTitle$(d As Long)
FormatDateForDayTitle$ = Format$(CDate(d), "dddd, mmm d, yyyy")
End Function

'EHT=None
Function FormatClientName(formatoption As ClientNameFormatMode, c As Client_DBPortion) As String
'Dim p&, lp&, t$, m As Boolean, matchingbrace$
Dim showname1 As Boolean, showname2 As Boolean

If Len(c.Person1.First) > 0 Then showname1 = True
If Len(c.Person2.First) > 0 Then showname2 = True
If showname1 And showname2 Then
    'If there are 2 people AND 1 is still living, then the deceased person is hidden
    'But if both people are deceased, show the names because tabSearch will cross out the whole list
    If (c.Person1.DOD = NullLong) Or (c.Person2.DOD = NullLong) Then
        If c.Person1.DOD <> NullLong Then
            showname1 = False
        ElseIf c.Person2.DOD <> NullLong Then
            showname2 = False
        End If
    End If
End If

Select Case formatoption
'johNSon, keNNeth (kEn) A & asHLeY (aSh) C [diVEr]
Case fSearchResults, fMailingList, fCustomListboxSorting, fChosenClients, fUnpaid, fSchedulePct
    FormatClientName = c.Person1.Last & ", " & FormatClientName(fExport_First, c)

'keNNeth (kEn) A & asHLeY (aSh) C [diVEr]
Case fExport_First
    If showname1 Then
        FormatClientName = FormatClientName & c.Person1.First
        If Len(c.Person1.Nickname) > 0 Then
            FormatClientName = FormatClientName & " (" & c.Person1.Nickname & ")"
        End If
        If Len(c.Person1.Initial) > 0 Then
            FormatClientName = FormatClientName & " " & UCase$(c.Person1.Initial)
        End If
    End If
    If showname2 Then
        If showname1 Then FormatClientName = FormatClientName & " & "
        FormatClientName = FormatClientName & c.Person2.First
        If Len(c.Person2.Nickname) > 0 Then
            FormatClientName = FormatClientName & " (" & c.Person2.Nickname & ")"
        End If
        If Len(c.Person2.Initial) > 0 Then
            FormatClientName = FormatClientName & " " & UCase$(c.Person2.Initial)
        End If
        If Len(c.Person2.Last) > 0 Then
            If c.Person2.Last <> c.Person1.Last Then
                FormatClientName = FormatClientName & " [" & c.Person2.Last & "]"
            End If
        End If
    End If

'johNSon
Case fExport_Last
    FormatClientName = c.Person1.Last

'keNNeth & asHLeY johNSon
Case fExport_Display
    If showname1 Then
        FormatClientName = c.Person1.First
    End If
    If showname2 Then
        If showname1 Then FormatClientName = FormatClientName & " & "
        FormatClientName = FormatClientName & c.Person2.First
    End If
    FormatClientName = FormatClientName & " " & c.Person1.Last

'KENNETH & ASHLEY JOHNSON
Case fPrintLabels
    If showname1 Then
        FormatClientName = UCase$(c.Person1.First)
    End If
    If showname2 Then
        If showname1 Then FormatClientName = FormatClientName & " & "
        FormatClientName = FormatClientName & UCase$(c.Person2.First)
    End If
    FormatClientName = FormatClientName & " " & UCase$(c.Person1.Last)

'JOHNSON,  keNNeth & asHLeY
Case fPullFiles
    FormatClientName = UCase$(c.Person1.Last) & ",  "       'Double-space is intentional, for display purposes!!!
    If showname1 Then
        FormatClientName = FormatClientName & c.Person1.First
    End If
    If showname2 Then
        If showname1 Then FormatClientName = FormatClientName & " & "
        FormatClientName = FormatClientName & c.Person2.First
    End If

'ClientID#1304[johNSon,keNNeth&asHLeY]
Case fLog
    FormatClientName = "ClientID#" & c.ID & "[" & c.Person1.Last & "," & c.Person1.First & IIf(c.Person2.First <> "", "&" & c.Person2.First, "") & "]"
End Select







'Old methods..........
'#########################################################
'    If Flag_IsSet(c.Flags, IncPtnrTrustEstate) Then
'        FormatClientName = c.Person1.Last & ", " & c.Person1.First
'    Else
'        Do
'            p = InStr(lp + 1, c.Person1.First, " ")
'            If p = 0 Then
'                If lp <= Len(c.Person1.First) Then
'                    p = Len(c.Person1.First) + 1
'                Else
'                    Exit Do
'                End If
'            End If
'            If p = (lp + 2) Then
'                If Mid$(c.Person1.First, p - 1, 1) = "&" Then FormatClientName = FormatClientName & "& "
'            Else
'                FormatClientName = FormatClientName & Mid$(c.Person1.First, lp + 1, p - lp - 1) & " "
'            End If
'            lp = p
'        Loop
'        FormatClientName = c.Person1.Last & ", " & Mid$(FormatClientName, 1, Len(FormatClientName) - 1)
'    End If
'#########################################################
'    FormatClientName = c.Person1.Last & ", "
'    If Len(c.Person1.First) > 0 Then
'        FormatClientName = FormatClientName & c.Person1.First
'        If Len(c.Person2.First) > 0 Then
'            FormatClientName = FormatClientName & " & " & c.Person2.First
'        End If
'    Else
'        FormatClientName = FormatClientName & c.Person2.First
'    End If
'#########################################################
'    If Flag_IsSet(c.Flags, IncPtnrTrustEstate) Then
'        FormatClientName = c.Person1.First
'    Else
'        'Remove parantheses and brackets
'        m = True
'        Do Until p > Len(c.Person1.First)
'            If m Then
'                p = p + 1
'                FormatClientName = Mid$(c.Person1.First, p, 1)
'                Select Case FormatClientName
'                Case "("
'                    matchingbrace$ = ")"
'                    m = False
'                Case "["
'                    matchingbrace$ = "]"
'                    m = False
'                Case Else
'                    t$ = t$ & FormatClientName
'                End Select
'            Else
'                p = InStr(lp + 1, c.Person1.First, matchingbrace$)
'                If p = 0 Then Exit Do
'                m = True
'            End If
'            lp = p
'        Loop
'
'        'Remove initials
'        lp = 0
'        Do
'            'Jump by spaces
'            p = InStr(lp + 1, t$, " ")
'            If p = 0 Then
'                If lp <= Len(t$) Then
'                    p = Len(t$) + 1
'                Else
'                    Exit Do
'                End If
'            End If
'            If p = (lp + 1) Then
'                'Double space, don't keep the second one
'            ElseIf p = (lp + 2) Then
'                'Only one letter, keep only if it's '&'
'                If Mid$(t$, p - 1, 1) = "&" Then FormatClientName = FormatClientName & "& "
'            Else
'                'Keep
'                FormatClientName = FormatClientName & Mid$(t$, lp + 1, p - lp - 1) & " "
'            End If
'            lp = p
'        Loop
'        FormatClientName = Mid$(FormatClientName, 1, Len(FormatClientName) - 1)
'    End If
'    If formatoption = fPrintLabels Then FormatClientName = UCase$(FormatClientName)
End Function

'EHT=None
Function FormatNumApptSlots$(nastu&)
Select Case nastu
Case 0:     FormatNumApptSlots$ = "-"
Case 1:     FormatNumApptSlots$ = "SA"
Case 2:     FormatNumApptSlots$ = "DA"
Case 3:     FormatNumApptSlots$ = "TA"
Case 4:     FormatNumApptSlots$ = "QA"
Case Else:  FormatNumApptSlots$ = nastu & "A"
End Select
End Function

'EHT=None
Function FormatRefDue$(p&)
If p < 0 Then
    FormatRefDue$ = "Due: " & Format$(-p, "$#,##0")
ElseIf p = 0 Then
    FormatRefDue$ = "$0"
Else
    FormatRefDue$ = "Ref: " & Format$(p, "$#,##0")
End If
End Function

'EHT=None
Function FormatTextForCSV(t$)
FormatTextForCSV = Replace$(t$, ",", ";")
End Function

'EHT=None
Function FormatTextForDB$(t$)
FormatTextForDB = Replace(Replace(Replace(Replace$(t$, _
                    vbTab, " "), _
                    vbCrLf, MultiLineSep), _
                    vbCr, MultiLineSep), _
                    vbLf, MultiLineSep)
End Function

'EHT=Standard
Function CalculateAge(dt1&, dt2&) As Long
On Error GoTo ERR_HANDLER

Dim m1&, d1&, y1&
Dim m2&, d2&, y2&
y1 = Year(dt1): m1 = Month(dt1): d1 = Day(dt1)
y2 = Year(dt2): m2 = Month(dt2): d2 = Day(dt2)
If m2 > m1 Then
    '2/28/2012 to 3/1/2015 = 3 yr old
    '2/29/2012 to 3/1/2015 = 3 yr old
    CalculateAge = y2 - y1
ElseIf m2 = m1 Then
    '2/28/2012 to 2/28/2015 = 3 yr old
    '2/29/2012 to 2/28/2015 = 2 yr old
    CalculateAge = (y2 - y1 - 1) - (d2 >= d1)   'Subtracting a boolean will add 1 if it's true
Else
    '2/28/2012 to 1/28/2015 = 2 yr old
    '2/29/2012 to 1/31/2015 = 2 yr old
    CalculateAge = (y2 - y1 - 1)
End If

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "CalculateAge", Err
End Function

'EHT=None
Function GetCapsLock() As Boolean
GetCapsLock = (GetKeyState(VK_CAPITAL) And 1 = 1)
End Function

'EHT=None
Function GetShiftState() As Integer
GetShiftState = (-1 * ((GetAsyncKeyState(VK_SHIFT) And &H8000&) = &H8000&)) Or _
                (-2 * ((GetAsyncKeyState(VK_CONTROL) And &H8000&) = &H8000&)) Or _
                (-4 * ((GetAsyncKeyState(VK_MENU) And &H8000&) = &H8000&))
End Function

'EHT=None
Sub Inc(ByRef t As Long)
t = t + 1
End Sub

'EHT=None
Sub IncBy(ByRef t As Long, i As Long)
If i = NullLong Then
    Err.Raise 1, , "Cannot increment by Null"
Else
    t = t + i
End If
End Sub

'EHT=None
Function IsLetterKey(KeyCode As Integer) As Boolean
IsLetterKey = ((KeyCode >= 65) And (KeyCode <= 90))
End Function

'EHT=None
Function IsNumberOrBlank(ByVal t$) As Boolean
If t$ = "" Then
    IsNumberOrBlank = True
Else
    t$ = Replace$(Replace$(t$, "$", ""), ",", "")
    IsNumberOrBlank = IsNumeric(t$)
End If
End Function

'EHT=None
Function JoinNumberArray1(SourceArray() As Long) As String
'Joins all
Dim a&, e As Boolean
For a = 0 To UBound(SourceArray)
    If e Then
        JoinNumberArray1 = JoinNumberArray1 & SEP1 & SourceArray(a)
    Else
        JoinNumberArray1 = JoinNumberArray1 & SourceArray(a)
        e = True
    End If
Next a
End Function

'EHT=None
Function JoinNumberArray2(SourceArray() As Long, SourceCount As Long) As String
'Joins only to SourceCount (allows for 0 elements)
Dim a&, e As Boolean
For a = 0 To SourceCount - 1
    If e Then
        JoinNumberArray2 = JoinNumberArray2 & SEP1 & SourceArray(a)
    Else
        JoinNumberArray2 = JoinNumberArray2 & SourceArray(a)
        e = True
    End If
Next a
End Function

'EHT=ResumeNext
Sub LostFocusFormat(txt As TextBox)
On Error Resume Next

Dim m As FieldFormatMode
m = Val(txt.Tag)
txt.Text = FieldToString(FieldFromString(txt.Text, m), m)
End Sub

'EHT=None
Sub MouseNullZone_Set(nzp&)
MouseNullZonePixels = nzp
GetCursorPos MouseNullZoneRef
End Sub

'EHT=None
Function MouseNullZone_Moved() As Boolean
Dim cp As POINTAPI
GetCursorPos cp
If Abs(cp.X - MouseNullZoneRef.X) > MouseNullZonePixels Then
    MouseNullZone_Moved = True
ElseIf Abs(cp.Y - MouseNullZoneRef.Y) > MouseNullZonePixels Then
    MouseNullZone_Moved = True
End If
End Function

'EHT=None
Sub PutKeyAsciiIntoTextbox(txt As TextBox, KeyAscii As Integer, ReplaceContents As Boolean)
Dim t$
t$ = Chr$(KeyAscii)
If ReplaceContents Then
    txt.Text = t$
Else
    txt.Text = txt.Text & t$
End If
txt.SelStart = Len(txt.Text)
txt.SelLength = 0
End Sub

'EHT=None
Sub PutKeyCodeIntoTextbox(txt As TextBox, KeyCode As Integer, ReplaceContents As Boolean)
Dim t$
If GetCapsLock Then
    t$ = Chr$(KeyCode)       'Uppercase
Else
    t$ = Chr$(KeyCode + 32)  'Lowercase
End If
If ReplaceContents Then
    txt.Text = t$
Else
    txt.Text = txt.Text & t$
End If
txt.SelStart = Len(txt.Text)
txt.SelLength = 0
End Sub

'EHT=Custom
Function RunningFromIDE() As Boolean
'Because debug statements are ignored when the app is compiled, the next statment will
'never be executed in the EXE. If we get an error then we are running in IDE / Debug mode
On Error GoTo e
Debug.Print 1 / 0
Exit Function
e: RunningFromIDE = True
End Function

'EHT=None
Sub SelectAll(txt As TextBox)
txt.SelStart = 0
txt.SelLength = Len(txt.Text)
End Sub

'EHT=None
Sub SelectFirstItemIfNoSelection(lst As Object)
If lst.ListIndex < 0 Then
    If lst.ListCount > 0 Then
        lst.ListIndex = 0
    End If
End If
End Sub

'EHT=None
Function SetTabStops(hwnd&, ParamArray TabStops()) As Boolean
'TabStops() is a 0 based array of tab stop values.
'The values "represent the number of quarters of the average character width for the font that is selected into the list box"
'For the standard VB6 listbox font, this comes out to be 3 pixels per 2 tabstop increments
'Also note that each tabstop is relative to the listbox, not relative to the previous tabstop
Dim ts&(), a%
ReDim ts(UBound(TabStops))
For a = 0 To UBound(TabStops)
    ts(a) = TabStops(a)
Next a
SetTabStops = (SendMessageByRef(hwnd, LB_SETTABSTOPS, UBound(TabStops) + 1, ts(0)) <> 0)
InvalidateRect hwnd, 0, 0
End Function

'EHT=ResumeNext
Sub HilightControl(frm As Form, ctrl As Object)
On Error Resume Next

Dim shp As Shape
Set shp = frm.Controls("hilight")
If Not shp Is Nothing Then
    If Not shp.Container Is ctrl.Container Then
        frm.Controls.Remove "hilight"
        Set shp = Nothing
    End If
End If
If shp Is Nothing Then
    Set shp = frm.Controls.Add("VB.Shape", "hilight", ctrl.Container)
    shp.BorderColor = &HC0&
    shp.BorderWidth = 2
    shp.ZOrder 0
End If
shp.Move ctrl.Left - 1, ctrl.Top - 1, ctrl.Width + 3, ctrl.Height + 3
shp.Visible = True
End Sub

'EHT=ResumeNext
Sub ClearControlHilight(frm As Form)
On Error Resume Next

Dim shp As Shape
Set shp = frm.Controls("hilight")
If Not shp Is Nothing Then
    shp.Visible = False
End If
End Sub

'EHT=ResumeNext
Sub SetControlTabOrder(frm As Form, taborder$)
On Error Resume Next

Dim ctrl As Object, tabordersplit$(), c$(), a&, focusset As Boolean
For Each ctrl In frm.Controls
    ctrl.TabStop = False
Next

tabordersplit = Split(taborder$, SEP1)
For a = 0 To UBound(tabordersplit)
    c$ = Split(tabordersplit(a), SEP2)
    Set ctrl = frm.Controls(c$(0))
    If UBound(c$) > 0 Then
        Set ctrl = ctrl(CInt(c$(1)))
    End If
    ctrl.TabIndex = a
    If TypeName(ctrl) = "OptionButton" Then
        If ctrl.Value Then
            ctrl.TabStop = True
            If Not focusset Then
                SetFocusWithoutErr ctrl
                focusset = True
            End If
        End If
    Else
        ctrl.TabStop = True
        If Not focusset Then
            SetFocusWithoutErr ctrl
            focusset = True
        End If
    End If
Next a
End Sub

'EHT=None
Function SplitNumberArray1(s$, a&()) As Long
'Splits string, redims a(), outputs to a(), returns count
Dim sp$(), b&
sp$ = Split(s$, SEP1)
ReDim a(UBound(sp$))
For b = 0 To UBound(sp$)
    a(b) = CLng(sp$(b))
Next b
SplitNumberArray1 = UBound(sp$) + 1
End Function

'EHT=None
Sub SplitNumberArray2(s$, a&())
'Splits string, outputs to a()
Dim sp$(), b&
sp$ = Split(s$, SEP1)
For b = 0 To UBound(sp$)
    a(b) = CLng(sp$(b))
Next b
End Sub

'EHT=ResumeNext
Sub TabToNextControl(frm As Form, selalliftextbox As Boolean, rev As Boolean)
On Error Resume Next

'Tab to the next control (or previus one if rev=True)
'If focus is on a control with TabStop = False, this code will find the next
' control in sequence that has TabStop = True
Dim o As Object, a&, ci&, cc&

ci = frm.ActiveControl.TabIndex
cc = frm.Controls.Count
If rev Then
    a = ci - 1
Else
    a = ci + 1
End If
Do
    If a < 0 Then a = cc - 1    'If at the beginning, wrap to the end
    If a >= cc Then a = 0       'If at the end, wrap to the beginning
    If a = ci Then Exit Do      'If wrapped back to current control, exit loop
    For Each o In frm.Controls
        If o.TabIndex = a Then
            If Not o.TabStop Then
                'This emply part must be here, because if the object doesn't have
                ' a TabStop property, the above line will error, causing the next line
                ' of code to be executed (this one)
            Else
                If o.Enabled Then
                    SetFocusWithoutErr o
                    If selalliftextbox Then
                        If TypeName(o) = "TextBox" Then SelectAll o
                    End If
                    Exit Do
                End If
            End If
        End If
    Next
    If rev Then
        a = a - 1
    Else
        a = a + 1
    End If
Loop
End Sub

'EHT=None
Function AddTrailingSlash(t$) As String
'If app is in root, returns c:\, but everywhere else, it does not include the '\'
'   Windows XP doesn't mind the '\\', but Windows 98 fails
If Right$(t$, 1) = "\" Then
    AddTrailingSlash = t$
Else
    AddTrailingSlash = t$ & "\"
End If
End Function

'EHT=Cleanup1
Function LoadGlobalSettings() As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

If GSLoaded Then Exit Function

Dim f$, fh As CMNMOD_CFileHandler
f$ = AppPath & App.EXEName & ".cfg"
'[Mark] Temp Code>
If FileExists(AppPath & App.EXEName & ".dat") And Not FileExists(f$) Then
    RenameFile AppPath & App.EXEName & ".dat", f$, True
End If
'<Temp Code
Set fh = OpenFile(f$, mBinary_Input)

GlobalSettings_Count = fh.ReadLong
If GlobalSettings_Count = 0 Then
    Erase GlobalSettings
Else
    ReDim GlobalSettings(GlobalSettings_Count - 1)
    Get #fh.FileNum, , GlobalSettings
End If

LoadGlobalSettings = True
GSLoaded = True

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "LoadGlobalSettings", Err, INCLEANUP: Resume CLEANUP
End Function

'EHT=Cleanup1
Function SaveGlobalSettings() As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean

If Not GSLoaded Then Exit Function
If Not GSChanged Then Exit Function

Dim fh As CMNMOD_CFileHandler, destfile$, tempfile$
destfile$ = AppPath & App.EXEName & ".cfg"
tempfile$ = destfile$ & ".sav"
Set fh = OpenFile(tempfile$, mBinary_Output)

fh.WriteLong GlobalSettings_Count
If GlobalSettings_Count > 0 Then
    Put #fh.FileNum, , GlobalSettings
End If
fh.CloseFile: Set fh = Nothing

RenameFile tempfile$, destfile$, True
GSChanged = False
SaveGlobalSettings = True

CLEANUP: INCLEANUP = True
    If Not fh Is Nothing Then fh.CloseFile: Set fh = Nothing

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SaveGlobalSettings", Err, INCLEANUP: Resume CLEANUP
End Function
