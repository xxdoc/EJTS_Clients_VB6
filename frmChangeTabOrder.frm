VERSION 5.00
Begin VB.Form frmChangeTabOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Tab Order"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstControls 
      Height          =   2415
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   8040
      Width           =   975
   End
   Begin VB.ListBox lstControls 
      Height          =   4335
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblControls 
      Caption         =   "(double-click to add to the top list)"
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
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label lblControls 
      Caption         =   "Other controls available:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label lblControls 
      Caption         =   "(use Ctrl+Up and Ctrl+Dn to reorder)"
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
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblControls 
      Caption         =   "Current tab-order:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmChangeTabOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "frmChangeTabOrder"

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again

Private ParentForm As Form
Private oldleft As Single

'EHT=Custom
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Sub Form_Show(frm As Form)
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

Dim c As Control, a%, i%, f As Boolean, t$, tabordersplit$()
Dim scr As RECT, par As RECT, ours As RECT, p1&, p2&
Dim offsetforaero As Long
Set ParentForm = frm

If Len(ParentForm.TabOrderSetting) = 0 Then Err.Raise 1, , "TabOrderSetting not initialized prior to calling frmChangeTabOrder"

t$ = DB_GetSetting(ActiveDBInstance, ParentForm.TabOrderSetting)
tabordersplit$ = Split(t$, SEP1)

'Populate the lists
lstControls(0).Clear
For a = 0 To UBound(tabordersplit$)
    lstControls(0).AddItem tabordersplit$(a)
Next a
lstControls(1).Clear
For Each c In ParentForm.Controls
    If IsControlTabable(c) Then
        'Determine the text to describe this control
        t$ = c.Name
        i = GetControlIndexWithoutError(c)
        If i >= 0 Then t$ = t$ & SEP2 & i

        'See if it is already in the tab order setting
        f = False
        For a = 0 To UBound(tabordersplit$)
            If tabordersplit$(a) = t$ Then f = True: Exit For
        Next a

        'If not found, then add it to the lower list
        If Not f Then lstControls(1).AddItem t$
    End If
Next
'If UBound(taborderSplit$) >= 0 Then
'    lstControls(0).ListIndex = 0       'This isn't working, so leave it out for now
'Else
    ClearControlHilight ParentForm
'End If

'Check if Win7 compositing is enabled. If it is, all requests (including API) for form position/size will be off by 5px
DwmIsCompositionEnabled offsetforaero
If offsetforaero Then offsetforaero = 5

'Calculate dimensions
SystemParametersInfo SPI_GETWORKAREA, 0, scr, 0
GetWindowRect ParentForm.hwnd, par
InflateRect par, offsetforaero, offsetforaero
GetWindowRect Me.hwnd, ours
InflateRect ours, offsetforaero, offsetforaero

'Position parent form
oldleft = ParentForm.Left
p1 = (par.Right + (ours.Right - ours.Left)) - scr.Right
If p1 > 0 Then
    If p1 > par.Left Then p1 = par.Left
    par.Left = par.Left - p1
    par.Right = par.Right - p1
    SetWindowPos ParentForm.hwnd, 0, par.Left + offsetforaero, par.Top + offsetforaero, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
End If

'Position our form
p1 = par.Right
If (p1 + (ours.Right - ours.Left)) > scr.Right Then p1 = scr.Right - (ours.Right - ours.Left)
p2 = par.Top
If (p2 + (ours.Bottom - ours.Top)) > scr.Bottom Then p2 = scr.Bottom - (ours.Bottom - ours.Top)
If p2 < scr.Top Then p2 = scr.Top
SetWindowPos Me.hwnd, 0, p1 + offsetforaero, p2 + offsetforaero, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
Me.Show 1, ParentForm

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Sub

'EHT=Standard
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ERR_HANDLER

ClearControlHilight ParentForm
ParentForm.Left = oldleft

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Unload", Err
End Sub

'EHT=Standard
Private Sub btnSave_Click()
On Error GoTo ERR_HANDLER

If Not btnSave.Enabled Then Exit Sub

Dim tabordersplit$(), taborder$, a&
ReDim tabordersplit$(lstControls(0).ListCount - 1)
For a = 0 To UBound(tabordersplit$)
    tabordersplit$(a) = lstControls(0).List(a)
Next a
taborder$ = Join(tabordersplit$, SEP1)
DB_SetSetting ActiveDBInstance, ParentForm.TabOrderSetting, taborder$

Me.Hide
SetControlTabOrder ParentForm, taborder$
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
Private Sub lstControls_Click(Index As Integer)
On Error GoTo ERR_HANDLER

If lstControls(Index).ListIndex >= 0 Then
    Dim ctrl As Control, c$, cs$()
    c$ = lstControls(Index).List(lstControls(Index).ListIndex)
    cs$ = Split(c$, SEP2)
    If UBound(cs$) > 0 Then
        Set ctrl = ParentForm.Controls(cs$(0))(CInt(cs$(1)))
    Else
        Set ctrl = ParentForm.Controls(cs$(0))
    End If
    HilightControl ParentForm, ctrl
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstControls_Click", Err
End Sub

'EHT=Standard
Private Sub lstControls_DblClick(Index As Integer)
On Error GoTo ERR_HANDLER

Dim i%, otherindex%
i = lstControls(Index).ListIndex
If i < 0 Then Exit Sub
otherindex = (Not (Index - 1)) + 1
'Add the item to the other listbox
lstControls(otherindex).AddItem lstControls(Index).List(i)
'Remove the item from this listbox
lstControls(Index).RemoveItem i
'Remove the hilight, since there is now no item selected
ClearControlHilight ParentForm

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstControls_DblClick", Err
End Sub

'EHT=Standard
Private Sub lstControls_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

'The user can only change the order of the items in the first listbox
If Index = 0 Then
    Dim i%, t$
    Select Case KeyCode
    Case vbKeyUp
        If Shift = vbCtrlMask Then
            i = lstControls(Index).ListIndex
            If i > 0 Then
                t$ = lstControls(Index).List(i)
                lstControls(Index).List(i) = lstControls(Index).List(i - 1)
                lstControls(Index).List(i - 1) = t$
            End If
        End If
    Case vbKeyDown
        If Shift = vbCtrlMask Then
            i = lstControls(Index).ListIndex
            If i >= 0 And i < (lstControls(Index).ListCount - 1) Then
                t$ = lstControls(Index).List(i)
                lstControls(Index).List(i) = lstControls(Index).List(i + 1)
                lstControls(Index).List(i + 1) = t$
            End If
        End If
    End Select
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstControls_KeyDown", Err
End Sub

'EHT=Standard
Private Sub lstControls_LostFocus(Index As Integer)
On Error GoTo ERR_HANDLER

lstControls(Index).ListIndex = -1
ClearControlHilight ParentForm

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lstControls_LostFocus", Err
End Sub

'EHT=Custom
Function GetControlIndexWithoutError(ctrl As Control) As Integer
On Error GoTo e
GetControlIndexWithoutError = -1
GetControlIndexWithoutError = ctrl.Index    'If succeeds, then this is a control array
Exit Function
e:
End Function

'EHT=Custom
Function GetControlTabIndexWithoutError(ctrl As Control) As Integer
On Error GoTo e
Dim ts As Boolean
GetControlTabIndexWithoutError = -2     'Assume the control cannot receive focus
If Not ctrl.Visible Then Exit Function
ts = ctrl.TabStop                       'If this errors, thin this control cannot receive focus

'OptionButtons are weird. Once one is selected, it's TabStop becomes True, and all others within the same Container becomes False
'This happens even if every OptionButton started with TabStop as False.
'There is no easy solution to this, so just assume all OptionButtons have TabStop True.
If TypeName(ctrl) = "OptionButton" Then ts = True

If ts Then
    GetControlTabIndexWithoutError = ctrl.TabIndex
Else
    GetControlTabIndexWithoutError = -1     'TabStop False
End If
Exit Function
e:
End Function

'EHT=Custom
Function IsControlTabable(ctrl As Control) As Boolean
On Error GoTo e
Dim ts As Boolean, ti As Integer
If Not ctrl.Visible Then Exit Function  'Control must be visible
ts = ctrl.TabStop                       'If this errors, thin this control cannot receive focus (the actual value doesn't matter)
ti = ctrl.TabIndex                      'If this errors, this would be weird
IsControlTabable = True
Exit Function
e:
End Function
