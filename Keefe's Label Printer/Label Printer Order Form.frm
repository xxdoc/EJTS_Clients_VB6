VERSION 5.00
Begin VB.Form frmOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Order"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "Label Printer Order Form.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4170
      TabIndex        =   6
      Top             =   3015
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2730
      TabIndex        =   5
      Top             =   3015
      Width           =   1215
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4980
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   855
      Width           =   405
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4980
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   405
   End
   Begin VB.ListBox lstNames 
      Height          =   2790
      ItemData        =   "Label Printer Order Form.frx":058A
      Left            =   120
      List            =   "Label Printer Order Form.frx":058C
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Name"
      Height          =   240
      Left            =   4965
      TabIndex        =   4
      Top             =   645
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Move"
      Height          =   240
      Left            =   4875
      TabIndex        =   3
      Top             =   450
      Width           =   615
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Canceled As Boolean

Private Sub cmdCancel_Click()
Hide
End Sub

Private Sub cmdMoveDown_Click()
X = lstNames.ListIndex
N$ = lstNames.List(X)
d = lstNames.ItemData(X)
lstNames.RemoveItem X
lstNames.AddItem N$, X + 1
lstNames.ItemData(X + 1) = d
lstNames.ListIndex = X + 1
End Sub

Private Sub cmdMoveUp_Click()
X = lstNames.ListIndex
N$ = lstNames.List(X)
d = lstNames.ItemData(X)
lstNames.RemoveItem X
lstNames.AddItem N$, X - 1
lstNames.ItemData(X - 1) = d
lstNames.ListIndex = X - 1
End Sub

Private Sub cmdOK_Click()
Canceled = False
Hide
End Sub

Private Sub lstNames_Click()
cmdMoveUp.Enabled = (lstNames.ListIndex > 0)
cmdMoveDown.Enabled = (lstNames.ListIndex < (lstNames.ListCount - 1))
End Sub

Private Sub lstNames_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 38 And cmdMoveUp.Enabled = True Then cmdMoveUp_Click: KeyCode = 0
If Shift = 2 And KeyCode = 40 And cmdMoveDown.Enabled = True Then cmdMoveDown_Click: KeyCode = 0
End Sub
