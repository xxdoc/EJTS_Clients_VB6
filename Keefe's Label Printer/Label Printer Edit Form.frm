VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "Label Printer Edit Form.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLabel 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   270
      MaxLength       =   33
      TabIndex        =   1
      Top             =   510
      Width           =   4470
   End
   Begin VB.TextBox txtLabel 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   270
      MaxLength       =   33
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   735
      Width           =   4470
   End
   Begin VB.TextBox txtLabel 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   270
      MaxLength       =   33
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   4470
   End
   Begin VB.TextBox txtLabel 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   270
      MaxLength       =   33
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1185
      Width           =   4470
   End
   Begin VB.TextBox txtLabel 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   270
      MaxLength       =   33
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1410
      Width           =   4470
   End
   Begin VB.TextBox txtLabel 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   270
      MaxLength       =   33
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1635
      Width           =   4470
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3645
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2205
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1395
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   4755
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1395
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   4755
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Canceled As Boolean

Private Sub cmdCancel_Click()
Hide
End Sub

Private Sub cmdOK_Click()
Canceled = False
Hide
End Sub

Private Sub txtLabel_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 38 And Index > 1 Then
    txtLabel(Index - 1).SetFocus
End If
If KeyCode = 40 And Index < 6 Then
    txtLabel(Index + 1).SetFocus
End If
End Sub

Private Sub txtLabel_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index < 6 Then
        txtLabel(Index + 1).SetFocus
    Else
        cmdOK.SetFocus
    End If
    KeyAscii = 0
End If
End Sub

