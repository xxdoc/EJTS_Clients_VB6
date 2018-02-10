VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   2640
   End
   Begin VB.ListBox lstControls 
      Height          =   4335
      IntegralHeight  =   0   'False
      Left            =   5280
      Sorted          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ParentForm As Form

Private Sub Command1_Click()
Command1.SetFocus
Unload Me
End Sub

Private Sub Form_Load()
Option1.TabStop = True
Option2.TabStop = True
Command1.TabStop = True
End Sub

Private Sub Text1_Change(Index As Integer)
Me.Caption = Me.Caption & "C"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Me.Caption = Me.Caption & "G"
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Me.Caption = Me.Caption & "L"
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Me.Caption = Me.Caption & "V"
If Text1(Index).Text = "" Then Cancel = True
End Sub

Private Sub Timer1_Timer()
Set ParentForm = Me

Dim c As Control, i%, ti%
lstControls.Clear
For Each c In ParentForm.Controls
    i = GetControlIndexWithoutError(c)
    ti = GetControlTabIndexWithoutError(c)
    'Stop
    'c.Parent
    lstControls.AddItem IIf(ti < 0, "", Format(ti, "000")) & vbTab & c.Name & IIf(i < 0, "", "," & i)
Next
End Sub

Function GetControlIndexWithoutError(ctrl As Control) As Integer
'CUSTOM ERROR HANDLING HERE INSTEAD OF TEMPLATE
''~~**##**~~**##**~~**##**~~**##**~~**##**~~**#

On Error GoTo e
GetControlIndexWithoutError = -1
GetControlIndexWithoutError = ctrl.Index
Exit Function
e:
End Function

Function GetControlTabIndexWithoutError(ctrl As Control) As Integer
'CUSTOM ERROR HANDLING HERE INSTEAD OF TEMPLATE
''~~**##**~~**##**~~**##**~~**##**~~**##**~~**#

On Error GoTo e
Dim ts As Boolean
GetControlTabIndexWithoutError = -1
ts = ctrl.TabStop
If Not ts Then
    If TypeName(ctrl) = "OptionButton" Then ts = True
End If
If ts Then GetControlTabIndexWithoutError = ctrl.TabIndex
Exit Function
e:
End Function
