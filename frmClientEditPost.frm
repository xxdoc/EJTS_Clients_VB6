VERSION 5.00
Begin VB.Form frmClientEditPost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit/Post Client"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12990
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientEditPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   866
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIncPtnrTrustEstate 
      Caption         =   "&INC, PTNR, TRUST, ESTATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton optInType 
      Caption         =   "&MI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   10320
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optInType 
      Caption         =   "&DO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   10320
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optInType 
      Caption         =   "&Appt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   10320
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optInType 
      Caption         =   "&NNTF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   11520
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkEFile 
      Caption         =   "E-Filed"
      Height          =   375
      Left            =   10320
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
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
      Index           =   136
      Left            =   5400
      TabIndex        =   116
      Tag             =   "31"
      Top             =   2640
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
      Index           =   135
      Left            =   5400
      TabIndex        =   114
      Tag             =   "31"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox pctFutureFlags 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   8520
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lblFlagNew_Title 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CurrYr:"
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
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   1095
      End
      Begin VB.Line Line8 
         X1              =   8
         X2              =   72
         Y1              =   40
         Y2              =   40
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inc"
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
         Left            =   120
         TabIndex        =   111
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comp"
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
         Left            =   120
         TabIndex        =   110
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NNTF"
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
         Index           =   5
         Left            =   120
         TabIndex        =   109
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ext"
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
         Index           =   6
         Left            =   120
         TabIndex        =   108
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Appt"
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
         Left            =   120
         TabIndex        =   107
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DO"
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
         Index           =   3
         Left            =   120
         TabIndex        =   106
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MI"
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
         Index           =   4
         Left            =   120
         TabIndex        =   105
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I/P/T/E"
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
         Index           =   7
         Left            =   120
         TabIndex        =   104
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New"
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
         Index           =   8
         Left            =   120
         TabIndex        =   103
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E-Filed"
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
         Index           =   9
         Left            =   120
         TabIndex        =   102
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblFlagFuture 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RBefPmt"
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
         Index           =   10
         Left            =   120
         TabIndex        =   101
         Top             =   3360
         Width           =   975
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3495
         Left            =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPctFlagsTitle 
         Caption         =   "Flags:"
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
         Left            =   0
         TabIndex        =   113
         Top             =   0
         Width           =   1095
      End
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
      Index           =   134
      Left            =   11640
      TabIndex        =   97
      TabStop         =   0   'False
      Tag             =   "21"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pctFlags 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   8040
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3960
      Width           =   2055
      Begin VB.Label lblFlagThisYear_Title 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CurrYr:"
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
         Left            =   1080
         TabIndex        =   95
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear_Title 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LastYr:"
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
         Left            =   120
         TabIndex        =   94
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inc"
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
         Left            =   1080
         TabIndex        =   93
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comp"
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
         Left            =   1080
         TabIndex        =   92
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NNTF"
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
         Index           =   5
         Left            =   1080
         TabIndex        =   91
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ext"
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
         Index           =   6
         Left            =   1080
         TabIndex        =   90
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Appt"
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
         Left            =   1080
         TabIndex        =   89
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DO"
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
         Index           =   3
         Left            =   1080
         TabIndex        =   88
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MI"
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
         Index           =   4
         Left            =   1080
         TabIndex        =   87
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I/P/T/E"
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
         Index           =   7
         Left            =   1080
         TabIndex        =   86
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New"
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
         Index           =   8
         Left            =   1080
         TabIndex        =   85
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E-Filed"
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
         Index           =   9
         Left            =   1080
         TabIndex        =   84
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inc"
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
         Left            =   120
         TabIndex        =   83
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comp"
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
         Left            =   120
         TabIndex        =   82
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NNTF"
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
         Index           =   5
         Left            =   120
         TabIndex        =   81
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ext"
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
         Index           =   6
         Left            =   120
         TabIndex        =   80
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Appt"
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
         Left            =   120
         TabIndex        =   79
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DO"
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
         Index           =   3
         Left            =   120
         TabIndex        =   78
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MI"
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
         Index           =   4
         Left            =   120
         TabIndex        =   77
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "I/P/T/E"
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
         Index           =   7
         Left            =   120
         TabIndex        =   76
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "New"
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
         Index           =   8
         Left            =   120
         TabIndex        =   75
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "E-Filed"
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
         Index           =   9
         Left            =   120
         TabIndex        =   74
         Top             =   3120
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   8
         X2              =   128
         Y1              =   40
         Y2              =   40
      End
      Begin VB.Label lblFlagLastYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RBefPmt"
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
         Index           =   10
         Left            =   120
         TabIndex        =   73
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lblFlagThisYear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RBefPmt"
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
         Index           =   10
         Left            =   1080
         TabIndex        =   72
         Top             =   3360
         Width           =   855
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3495
         Left            =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblPctFlagsTitle 
         Caption         =   "Flags (click to toggle):"
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
         Index           =   0
         Left            =   0
         TabIndex        =   96
         Top             =   0
         Width           =   1935
      End
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
      Index           =   133
      Left            =   5640
      TabIndex        =   34
      Tag             =   "70"
      Top             =   4680
      Width           =   975
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
      Index           =   132
      Left            =   4560
      TabIndex        =   33
      Tag             =   "70"
      Top             =   4680
      Width           =   975
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
      Index           =   113
      Left            =   6720
      TabIndex        =   13
      Tag             =   "31"
      Top             =   2640
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
      Index           =   109
      Left            =   5880
      TabIndex        =   10
      Tag             =   "51"
      Top             =   1920
      Width           =   495
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
      Index           =   108
      Left            =   4080
      TabIndex        =   9
      Tag             =   "50"
      Top             =   1920
      Width           =   1695
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
      Index           =   110
      Left            =   6480
      TabIndex        =   11
      Tag             =   "50"
      Top             =   1920
      Width           =   3615
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
      Index           =   107
      Left            =   1440
      TabIndex        =   8
      Tag             =   "50"
      Top             =   1920
      Width           =   2535
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
      Index           =   106
      Left            =   6720
      TabIndex        =   6
      Tag             =   "31"
      Top             =   1080
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
      Index           =   102
      Left            =   5880
      TabIndex        =   3
      Tag             =   "51"
      Top             =   360
      Width           =   495
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
      Index           =   101
      Left            =   4080
      TabIndex        =   2
      Tag             =   "50"
      Top             =   360
      Width           =   1695
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
      Index           =   112
      Left            =   1440
      TabIndex        =   12
      Tag             =   "52"
      Top             =   2640
      Width           =   3855
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
      Index           =   122
      Left            =   11640
      TabIndex        =   23
      Tag             =   "23"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      Height          =   855
      Index           =   125
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   "53"
      Top             =   6360
      Width           =   6495
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
      Index           =   128
      Left            =   10320
      TabIndex        =   29
      Tag             =   "21"
      Top             =   4080
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
      Index           =   129
      Left            =   10320
      TabIndex        =   30
      Tag             =   "21"
      Top             =   5520
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
      Index           =   130
      Left            =   10320
      TabIndex        =   31
      Tag             =   "21"
      Top             =   3360
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
      Index           =   127
      Left            =   10320
      TabIndex        =   28
      Tag             =   "23"
      Top             =   6960
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
      Index           =   126
      Left            =   10320
      TabIndex        =   27
      Tag             =   "23"
      Top             =   6240
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
      Index           =   131
      Left            =   10320
      TabIndex        =   32
      Tag             =   "54"
      Top             =   4800
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
      Index           =   123
      Left            =   10320
      TabIndex        =   24
      Tag             =   "31"
      Top             =   1440
      Width           =   1815
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
      Index           =   121
      Left            =   11280
      TabIndex        =   22
      Tag             =   "12"
      Top             =   2160
      Width           =   855
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
      Index           =   124
      Left            =   10320
      TabIndex        =   25
      Tag             =   "12"
      Top             =   2160
      Width           =   855
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
      Index           =   120
      Left            =   3240
      TabIndex        =   21
      Tag             =   "13"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Height          =   615
      Index           =   118
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Tag             =   "50"
      Top             =   5400
      Width           =   6495
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
      Index           =   105
      Left            =   1440
      TabIndex        =   5
      Tag             =   "52"
      Top             =   1080
      Width           =   3855
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
      Index           =   117
      Left            =   6480
      TabIndex        =   18
      Tag             =   "51"
      Top             =   3960
      Width           =   1455
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
      Index           =   116
      Left            =   5880
      TabIndex        =   17
      Tag             =   "51"
      Top             =   3960
      Width           =   495
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
      Index           =   115
      Left            =   1440
      TabIndex        =   16
      Tag             =   "51"
      Top             =   3960
      Width           =   4335
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
      Index           =   114
      Left            =   1440
      TabIndex        =   15
      Tag             =   "51"
      Top             =   3480
      Width           =   6495
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
      Index           =   111
      Left            =   8040
      TabIndex        =   14
      Tag             =   "60"
      Top             =   2640
      Width           =   2055
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
      Index           =   104
      Left            =   8040
      TabIndex        =   7
      Tag             =   "60"
      Top             =   1080
      Width           =   2055
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
      Index           =   119
      Left            =   8040
      TabIndex        =   20
      Tag             =   "60"
      Top             =   3480
      Width           =   2055
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
      Index           =   100
      Left            =   1440
      TabIndex        =   1
      Tag             =   "50"
      Top             =   360
      Width           =   2535
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
      Index           =   103
      Left            =   6480
      TabIndex        =   4
      Tag             =   "50"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5880
      TabIndex        =   36
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton btnSavePost 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   35
      Top             =   7320
      Width           =   3135
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      Caption         =   "114 yr old today"
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
      Left            =   5880
      TabIndex        =   119
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      Caption         =   "114 yr old today"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000D0&
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   118
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "DOB:"
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
      Index           =   136
      Left            =   5400
      TabIndex        =   117
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "DOB:"
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
      Index           =   135
      Left            =   5400
      TabIndex        =   115
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "State returns (if any):"
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
      Index           =   23
      Left            =   10320
      TabIndex        =   69
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "St result 2:"
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
      Index           =   134
      Left            =   11640
      TabIndex        =   98
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Newest:"
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
      Index           =   36
      Left            =   5640
      TabIndex        =   71
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Oldest:"
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
      Index           =   35
      Left            =   4560
      TabIndex        =   70
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "Common"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   120
      TabIndex        =   56
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Cell phone"
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
      Index           =   34
      Left            =   8040
      TabIndex        =   53
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "DOD:"
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
      Index           =   113
      Left            =   7440
      TabIndex        =   52
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Initial:"
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
      Index           =   32
      Left            =   5880
      TabIndex        =   49
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Nickname:"
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
      Index           =   31
      Left            =   4080
      TabIndex        =   48
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Email address:"
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
      Index           =   30
      Left            =   1440
      TabIndex        =   51
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label lbl 
      Caption         =   "Last name (if different than taxpayer's):"
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
      Index           =   29
      Left            =   6480
      TabIndex        =   47
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label lbl 
      Caption         =   "First name:"
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
      Index           =   11
      Left            =   1440
      TabIndex        =   46
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label lbl 
      Caption         =   "Cell phone"
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
      Index           =   10
      Left            =   8040
      TabIndex        =   44
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "DOD:"
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
      Index           =   106
      Left            =   7440
      TabIndex        =   45
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Initial:"
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
      Left            =   5880
      TabIndex        =   39
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Nickname:"
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
      Index           =   2
      Left            =   4080
      TabIndex        =   40
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "LY Fee:"
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
      Index           =   14
      Left            =   11640
      TabIndex        =   60
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Operation notes:"
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
      Index           =   125
      Left            =   1440
      TabIndex        =   65
      Top             =   6120
      Width           =   6495
   End
   Begin VB.Label lbl 
      Caption         =   "First name:"
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
      Index           =   0
      Left            =   1440
      TabIndex        =   38
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lbl 
      Caption         =   "Last name:"
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
      Left            =   6480
      TabIndex        =   41
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lbl 
      Caption         =   "St result:"
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
      Index           =   22
      Left            =   10320
      TabIndex        =   68
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Fed result:"
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
      Index           =   21
      Left            =   10320
      TabIndex        =   67
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "AGI:"
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
      Index           =   20
      Left            =   10320
      TabIndex        =   66
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "$ Owed:"
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
      Index           =   19
      Left            =   10320
      TabIndex        =   64
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Prep fee:"
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
      Index           =   18
      Left            =   10320
      TabIndex        =   63
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Minutes:"
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
      Index           =   16
      Left            =   10320
      TabIndex        =   62
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "Completed on:"
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
      Index           =   15
      Left            =   10320
      TabIndex        =   61
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "LY Min:"
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
      Index           =   13
      Left            =   11280
      TabIndex        =   59
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lbl 
      Caption         =   "# Slots:"
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
      Index           =   12
      Left            =   3240
      TabIndex        =   58
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label lbl 
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
      Index           =   6
      Left            =   1440
      TabIndex        =   57
      Top             =   5160
      Width           =   6495
   End
   Begin VB.Label lbl 
      Caption         =   "Email address:"
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
      Index           =   5
      Left            =   1440
      TabIndex        =   43
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lbl 
      Caption         =   "Home phone"
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
      Index           =   8
      Left            =   8040
      TabIndex        =   55
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lbl 
      Caption         =   "Address:"
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
      Index           =   4
      Left            =   1440
      TabIndex        =   54
      Top             =   3240
      Width           =   6495
   End
   Begin VB.Label lbl 
      Caption         =   "Spouse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   50
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lbl 
      Caption         =   "Taxpayer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   120
      TabIndex        =   42
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   680
      X2              =   680
      Y1              =   8
      Y2              =   488
   End
   Begin VB.Line Line2 
      X1              =   16
      X2              =   672
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line3 
      X1              =   16
      X2              =   672
      Y1              =   208
      Y2              =   208
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
      Left            =   11880
      TabIndex        =   99
      Top             =   8040
      Width           =   975
   End
End
Attribute VB_Name = "frmClientEditPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MOD_NAME = "frmClientEditPost"

Private FormLoadedAlready As Boolean        'Safety variable to ensure all references to this form are erased before attempting to load it again
Public TabOrderSetting As String            'This is set in Form_Show, depending on the post/edit mode

Enum enumShowFormMode
    fPost
    fEdit
    fNew
End Enum

Private Enum FieldName
    fPerson1First = 100
    fPerson1Nickname
    fPerson1Initial
    fPerson1Last
    fPerson1Phone
    fPerson1Email
    fPerson1DOD

    fPerson2First
    fPerson2Nickname
    fPerson2Initial
    fPerson2Last
    fPerson2Phone
    fPerson2Email
    fPerson2DOD

    fAddressStreet
    fAddressCity
    fAddressState
    fAddressZipCode
    fNotes

    fPhoneHome
    fNumApptSlotsToUse
    fLastYear_MinutesToComplete
    fLastYear_PrepFee
    fCompletionDate
    fMinutesToComplete

    fOperationNotes

    fPrepFee
    fMoneyOwed
    fResultFederal
    fResultState
    fResultAGI
    fStateList

    fOldestYearFiled
    fNewestYearFiled

    fResultState2

    fPerson1DOB
    fPerson2DOB
End Enum

Private ShowFormMode As enumShowFormMode
Private PreviouslyMarkedIncomplete As Boolean
Private DontChangeFlags As Boolean
Private DontChangeFocus As Long     'Was originally a boolean, but now is a counter of the pieces of code which need this set to True
Private tempclient As Client
Private thisID&
Private Changed As Boolean

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Function Form_Show(cID&, mShowFormMode As enumShowFormMode, Optional OwnerForm_OtherThanFrmMainOrTabs As Form, Optional NewClientInputString As String) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

'If Post/Edit, cID(input) is the ClientID to open
'If New, cID(output) is where the new ClientID is returned, or -1 if canceled

Dim cindex&, a&, f&

ShowFormMode = mShowFormMode
If OwnerForm_OtherThanFrmMainOrTabs Is Nothing Then Set OwnerForm_OtherThanFrmMainOrTabs = frmMain
DontChangeFocus = 1

If ShowFormMode = fNew Then
    thisID = -1
Else
    thisID = cID
    cindex = DB_FindClientIndex(ActiveDBInstance, cID)
    If cindex < 0 Then
        Err.Raise 1, , "Client #" & cID & " not found!"
    End If
    tempclient = ActiveDBInstance.Clients(cindex)
End If

If ShowFormMode = fPost Then
    chkIncPtnrTrustEstate.Visible = True
    optInType(0).Visible = True
    optInType(1).Visible = True
    optInType(2).Visible = True
    optInType(3).Visible = True
    chkEFile.Visible = True
    txtField(fResultState2).Visible = True
    lbl(fResultState2).Visible = True
    pctFlags.Visible = False
    pctFutureFlags.Visible = True
End If

With tempclient.c
    If ShowFormMode = fNew Then
        'Initialize a few fields that are not defaulted to the correct values by Visual Basic
        If NewClientInputString$ <> "" Then
            a = InStr(NewClientInputString$, ",")
            If a Then
                .Person1.Last = CapatalizeFirstLetter(Trim$(Left$(NewClientInputString$, a - 1)))
                .Person1.First = CapatalizeFirstLetter(Trim$(Mid$(NewClientInputString$, a + 1)))
            Else
                .Person1.Last = CapatalizeFirstLetter(Trim$(NewClientInputString$))
            End If
        End If
        .Person1.DOB = NullLong
        .Person1.DOD = NullLong
        .Person2.DOB = NullLong
        .Person2.DOD = NullLong

        .NumApptSlotsToUse = 0  'New clients should be set to Auto (CHOS will calculate it properly later)
        .Flags = NewClient

        .LastYear_MinutesToComplete = NullLong
        .LastYear_PrepFee = NullLong
        .LastYear_Flags = 0
        .OldestYearFiled = NullLong
        .NewestYearFiled = NullLong

        .CompletionDate = NullLong
        .MinutesToComplete = NullLong
        .PrepFee = NullLong
        .MoneyOwed = NullLong
        .ResultAGI = NullLong
        .ResultFederal = NullLong
        .ResultState = NullLong
    End If

    FieldToTextbox txtField(fPerson1First), .Person1.First, True
    FieldToTextbox txtField(fPerson1Nickname), .Person1.Nickname, True
    FieldToTextbox txtField(fPerson1Initial), .Person1.Initial, True
    FieldToTextbox txtField(fPerson1Last), .Person1.Last, True
    FieldToTextbox txtField(fPerson1Phone), .Person1.Phone, True
    FieldToTextbox txtField(fPerson1Email), .Person1.Email, True
    FieldToTextbox txtField(fPerson1DOB), .Person1.DOB, True
    FieldToTextbox txtField(fPerson1DOD), .Person1.DOD, True

    FieldToTextbox txtField(fPerson2First), .Person2.First, True
    FieldToTextbox txtField(fPerson2Nickname), .Person2.Nickname, True
    FieldToTextbox txtField(fPerson2Initial), .Person2.Initial, True
    FieldToTextbox txtField(fPerson2Last), .Person2.Last, True
    FieldToTextbox txtField(fPerson2Phone), .Person2.Phone, True
    FieldToTextbox txtField(fPerson2Email), .Person2.Email, True
    FieldToTextbox txtField(fPerson2DOB), .Person2.DOB, True
    FieldToTextbox txtField(fPerson2DOD), .Person2.DOD, True

    FieldToTextbox txtField(fAddressStreet), .AddressStreet, True
    FieldToTextbox txtField(fAddressCity), .AddressCity, True
    FieldToTextbox txtField(fAddressState), .AddressState, True
    FieldToTextbox txtField(fAddressZipCode), .AddressZipCode, True
    FieldToTextbox txtField(fNotes), .Notes, (ShowFormMode <> fPost)

    FieldToTextbox txtField(fPhoneHome), .PhoneHome, True
    FieldToTextbox txtField(fNumApptSlotsToUse), .NumApptSlotsToUse, (ShowFormMode <> fPost)
    FieldToTextbox txtField(fLastYear_MinutesToComplete), .LastYear_MinutesToComplete, (ShowFormMode <> fPost)
    FieldToTextbox txtField(fLastYear_PrepFee), .LastYear_PrepFee, (ShowFormMode <> fPost)
    If (ShowFormMode = fPost) And (.CompletionDate = NullLong) Then
        FieldToTextbox txtField(fCompletionDate), Date, True
    Else
        FieldToTextbox txtField(fCompletionDate), .CompletionDate, True
    End If
    FieldToTextbox txtField(fMinutesToComplete), .MinutesToComplete, True

    FieldToTextbox txtField(fOperationNotes), .OpNotes, (ShowFormMode <> fPost)

    FieldToTextbox txtField(fPrepFee), .PrepFee, True
    FieldToTextbox txtField(fMoneyOwed), .MoneyOwed, True
    FieldToTextbox txtField(fResultAGI), .ResultAGI, True
    FieldToTextbox txtField(fResultFederal), .ResultFederal, True
    If (ShowFormMode = fPost) And (Len(.StateList) = 0) Then
        FieldToTextbox txtField(fStateList), DB_GetSetting(ActiveDBInstance, "GLOBAL_DefaultState"), True
    Else
        FieldToTextbox txtField(fStateList), .StateList, True
    End If
    FieldToTextbox txtField(fResultState), .ResultState, True
    FieldToTextbox txtField(fResultState2), NullLong, (ShowFormMode = fPost)

    FieldToTextbox txtField(fOldestYearFiled), .OldestYearFiled, (ShowFormMode <> fPost)
    FieldToTextbox txtField(fNewestYearFiled), .NewestYearFiled, (ShowFormMode <> fPost)

    pctFlags.Enabled = (ShowFormMode <> fPost)
    If ShowFormMode = fPost Then
        For a = 0 To ClientFlags_DATAITEMUBOUND
            SetFutureFlagIndicator False, 0, a
        Next a

        'Make a working copy of the Flags
        PreviouslyMarkedIncomplete = Flag_IsSet(.Flags, PartiallyComplete)

        'Initialize them for posting...
        '   Inc       '0    OFF always
        'x  Comp      '1    ON if not NNTF
        'x  Appt      '?    ON if not DO/MI/NNTF
        'x  DO        'same Copy from DB
        'x  MI        'same Copy from DB
        'x  NNTF      '?    ON if NNTF LY
        'x  Ext       'same Copy from DB
        '   I/P/T/E   '?    ON if IPTE LY (takes precedence over NNTF, since you can't have both)
        'x  New       'same Copy from DB
        '   E-Filed   '?    ON if Not IPTE
        'x  RBefPmt   'same Copy from DB

        Me.Visible = True

        'Copy over the 'same' flags (above)
        SetFutureFlagIndicator Flag_IsSet(.Flags, Extension), Extension
        SetFutureFlagIndicator Flag_IsSet(.Flags, NewClient), NewClient
        SetFutureFlagIndicator Flag_IsSet(.Flags, ReleasedBeforePayment), ReleasedBeforePayment


        'Then initialize a few of them, depending on certain conditions
        chkEFile.Value = vbChecked  'Everything is E-filed by default
        If Flag_IsSet(.LastYear_Flags, IncPtnrTrustEstate) Then
            chkIncPtnrTrustEstate.Value = vbChecked
        End If
        If Flag_IsSet(.Flags, DroppedOff) Then
            optInType(1).Value = True
        ElseIf Flag_IsSet(.Flags, MailedIn) Then
            optInType(2).Value = True
        ElseIf Flag_IsSet(.LastYear_Flags, NoNeedToFile) And (Not Flag_IsSet(.LastYear_Flags, IncPtnrTrustEstate)) Then
            optInType(3).Value = True
        Else
            optInType(0).Value = True
        End If
        If IsFutureFlagIndicatorSet(EFiled) And (Not Flag_IsSet(.LastYear_Flags, IncPtnrTrustEstate)) Then
            chkEFile.Value = vbChecked
        End If
    Else
        For a = 0 To ClientFlags_DATAITEMUBOUND
            f = 2 ^ a
            SetCYFlagIndicator Flag_IsSet(.Flags, f), 0, a
            SetLYFlagIndicator Flag_IsSet(.LastYear_Flags, f), 0, a
        Next a
    End If
End With
If ShowFormMode = fNew Then
    Me.Caption = "New Client"
Else
    Me.Caption = "Client #" & tempclient.c.ID & " - " & Choose(ShowFormMode + 1, "Post", "Edit")
End If
UpdateDOBandDODtext
btnSavePost.Caption = IIf(ShowFormMode = fPost, "&Post", "Save")
btnSavePost.Enabled = ActiveDBInstance.IsWriteable
If ShowFormMode = fPost Then
    'Set a new tab order for Post Mode
    TabOrderSetting = "GLOBAL_TabOrder_ClientPost"
    SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)
Else
    'Set a new tab order for Edit/New Mode
    TabOrderSetting = "GLOBAL_TabOrder_ClientEdit"
    SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)
End If
frmMain.IdlePauseTimeout
DontChangeFocus = 0
Changed = False
'-----------------------------------
Me.Visible = False
Me.Show 1, OwnerForm_OtherThanFrmMainOrTabs
'-----------------------------------
Form_Show = Changed
frmMain.IdleSetAction
If ShowFormMode = fNew Then cID = thisID

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Function

'EHT=Standard
Public Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ERR_HANDLER

Select Case KeyAscii
Case vbKeyReturn
    If Not Me.ActiveControl Is txtField(fOperationNotes) Then
        KeyAscii = 0    'Stop the beep
    End If
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyPress", Err
End Sub

'EHT=Standard
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

Select Case KeyCode
Case vbKeyReturn
    If Shift = vbCtrlMask Then
        If DontChangeFocus = 0 Then SetFocusWithoutErr btnSavePost
        btnSavePost_Click
    ElseIf Not Me.ActiveControl Is txtField(fOperationNotes) Then
        TabToNextControl Me, True, (Shift = vbShiftMask)
    End If
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=Standard
Private Sub btnSavePost_Click()
On Error GoTo ERR_HANDLER

If Not btnSavePost.Enabled Then Exit Sub

Dim cindex&, a&, p1&, p2&

If ShowFormMode <> fNew Then
    cindex = DB_FindClientIndex(ActiveDBInstance, thisID)
    If cindex < 0 Then
        Err.Raise 1, , "Unable to save. Client #" & thisID & " not found!"
    End If
    'tempclient = ActiveDBInstance.Clients(cindex)  'Make a temp copy
End If

With tempclient.c
    If (ShowFormMode = fPost) Then
        .Flags = 0
        For a = 0 To ClientFlags_DATAITEMUBOUND
            If IsFutureFlagIndicatorSet(0, a) Then
                .Flags = .Flags Or (2 ^ a)
            End If
        Next a
    Else
        .Flags = 0
        .LastYear_Flags = 0
        For a = 0 To ClientFlags_DATAITEMUBOUND
            If IsCYFlagIndicatorSet(0, a) Then
                .Flags = .Flags Or (2 ^ a)
            End If
            If IsLYFlagIndicatorSet(0, a) Then
                .LastYear_Flags = .LastYear_Flags Or (2 ^ a)
            End If
        Next a
    End If

    FieldFromTextbox txtField(fPerson1First), .Person1.First
    FieldFromTextbox txtField(fPerson1Nickname), .Person1.Nickname
    FieldFromTextbox txtField(fPerson1Initial), .Person1.Initial
    FieldFromTextbox txtField(fPerson1Last), .Person1.Last
    FieldFromTextbox txtField(fPerson1Phone), .Person1.Phone
    FieldFromTextbox txtField(fPerson1Email), .Person1.Email
    FieldFromTextbox txtField(fPerson1DOB), .Person1.DOB
    FieldFromTextbox txtField(fPerson1DOD), .Person1.DOD

    FieldFromTextbox txtField(fPerson2First), .Person2.First
    FieldFromTextbox txtField(fPerson2Nickname), .Person2.Nickname
    FieldFromTextbox txtField(fPerson2Initial), .Person2.Initial
    FieldFromTextbox txtField(fPerson2Last), .Person2.Last
    FieldFromTextbox txtField(fPerson2Phone), .Person2.Phone
    FieldFromTextbox txtField(fPerson2Email), .Person2.Email
    FieldFromTextbox txtField(fPerson2DOB), .Person2.DOB
    FieldFromTextbox txtField(fPerson2DOD), .Person2.DOD

    FieldFromTextbox txtField(fAddressStreet), .AddressStreet
    FieldFromTextbox txtField(fAddressCity), .AddressCity
    FieldFromTextbox txtField(fAddressState), .AddressState
    FieldFromTextbox txtField(fAddressZipCode), .AddressZipCode
    FieldFromTextbox txtField(fNotes), .Notes

    FieldFromTextbox txtField(fPhoneHome), .PhoneHome
    FieldFromTextbox txtField(fNumApptSlotsToUse), .NumApptSlotsToUse
    FieldFromTextbox txtField(fLastYear_MinutesToComplete), .LastYear_MinutesToComplete
    FieldFromTextbox txtField(fLastYear_PrepFee), .LastYear_PrepFee
    If Flag_IsSet(.Flags, NoNeedToFile) Then
        .CompletionDate = NullLong
    Else
        FieldFromTextbox txtField(fCompletionDate), .CompletionDate
    End If
    FieldFromTextbox txtField(fMinutesToComplete), .MinutesToComplete

    FieldFromTextbox txtField(fOperationNotes), .OpNotes

    FieldFromTextbox txtField(fPrepFee), .PrepFee
    FieldFromTextbox txtField(fMoneyOwed), .MoneyOwed
    FieldFromTextbox txtField(fResultAGI), .ResultAGI
    FieldFromTextbox txtField(fResultFederal), .ResultFederal
    FieldFromTextbox txtField(fStateList), .StateList
    If (ShowFormMode = fPost) Then
        FieldFromTextbox txtField(fResultState), p1
        FieldFromTextbox txtField(fResultState2), p2
        If Len(.StateList) = 0 Then
            If (p1 <> NullLong) Or (p2 <> NullLong) Then
                ShowErrorMsg "If no states are listed, then there cannot be any state results entered either."
                Exit Sub
            End If
            .ResultState = NullLong
        Else
            'The user can enter the single value in either box. We'll just shift it over to p1
            If (p1 = NullLong) And (p2 <> NullLong) Then p1 = p2: p2 = NullLong
            If Len(.StateList) = 2 Then
                If p1 = NullLong Then
                    ShowErrorMsg "If a state is listed, there must also be a state result entered."
                    Exit Sub
                End If
                If p2 <> NullLong Then
                    ShowErrorMsg "Only one state is listed, yet there are two results entered."
                    Exit Sub
                End If
                .ResultState = p1
            Else
                If (p1 = NullLong) Or (p2 = NullLong) Then
                    ShowErrorMsg "If two or more states are listed, there must be two state results entered."
                    Exit Sub
                End If
                .ResultState = p1 + p2
            End If
        End If
    Else
        FieldFromTextbox txtField(fResultState), .ResultState
    End If
    FieldFromTextbox txtField(fOldestYearFiled), .OldestYearFiled
    FieldFromTextbox txtField(fNewestYearFiled), .NewestYearFiled

    If ShowFormMode = fNew Then
        .ID = DB_GetNewClientID(ActiveDBInstance)
        DB_AddClient ActiveDBInstance, tempclient
        frmMain.SetChangedFlagAndIndication
        tabLogFile.WriteLine "Created " & FormatClientName(fLog, tempclient.c)
        thisID = .ID
    Else
        'Save temp copy back to database
        tempclient.Temp_RegenerateTempData = True
        ActiveDBInstance.Clients(cindex) = tempclient
        frmMain.SetChangedFlagAndIndication
        tabLogFile.WriteLine "Edited " & FormatClientName(fLog, tempclient.c)
    End If
End With

Changed = True
Unload Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSavePost_Click", Err
End Sub

'EHT=Standard
Private Sub btnSavePost_GotFocus()
On Error GoTo ERR_HANDLER

HilightControl Me, btnSavePost

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSavePost_GotFocus", Err
End Sub

'EHT=Standard
Private Sub btnSavePost_LostFocus()
On Error GoTo ERR_HANDLER

ClearControlHilight Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSavePost_LostFocus", Err
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
Private Sub btnCancel_GotFocus()
On Error GoTo ERR_HANDLER

HilightControl Me, btnCancel

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnCancel_GotFocus", Err
End Sub

'EHT=Standard
Private Sub btnCancel_LostFocus()
On Error GoTo ERR_HANDLER

ClearControlHilight Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnCancel_LostFocus", Err
End Sub

'EHT=Standard
Private Sub chkEFile_Click()
On Error GoTo ERR_HANDLER

If Not DontChangeFlags Then SetFutureFlagIndicator (chkEFile.Value = vbChecked), EFiled
'If DontChangeFocus = 0 Then SetFocusWithoutErr txtField(fResultAGI)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkEFile_Click", Err
End Sub

'EHT=Standard
Private Sub chkEFile_GotFocus()
On Error GoTo ERR_HANDLER

HilightControl Me, chkEFile

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkEFile_GotFocus", Err
End Sub

'EHT=Standard
Private Sub chkEFile_LostFocus()
On Error GoTo ERR_HANDLER

ClearControlHilight Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkEFile_LostFocus", Err
End Sub

'EHT=Standard
Private Sub chkIncPtnrTrustEstate_Click()
On Error GoTo ERR_HANDLER

If Not chkIncPtnrTrustEstate.Enabled Then Exit Sub

Dim reg As Boolean
reg = (chkIncPtnrTrustEstate.Value = vbUnchecked)

If reg Then
    chkIncPtnrTrustEstate.BackColor = vbButtonFace
Else
    chkIncPtnrTrustEstate.BackColor = &HC0E0FF
End If

If Not DontChangeFlags Then SetFutureFlagIndicator (Not reg), IncPtnrTrustEstate

optInType(3).Enabled = reg
EnableTextbox txtField(fResultFederal), reg
EnableTextbox txtField(fResultState), reg
EnableTextbox txtField(fResultState2), reg
EnableTextbox txtField(fStateList), reg
EnableTextbox txtField(fResultAGI), reg
EnableTextbox txtField(fPerson1Nickname), reg
EnableTextbox txtField(fPerson1Initial), reg
EnableTextbox txtField(fPerson1DOB), reg
EnableTextbox txtField(fPerson1DOD), reg
EnableTextbox txtField(fPerson2First), reg
EnableTextbox txtField(fPerson2Nickname), reg
EnableTextbox txtField(fPerson2Initial), reg
EnableTextbox txtField(fPerson2Last), reg
EnableTextbox txtField(fPerson2Email), reg
EnableTextbox txtField(fPerson2DOB), reg
EnableTextbox txtField(fPerson2DOD), reg
EnableTextbox txtField(fPerson2Phone), reg
EnableTextbox txtField(fPhoneHome), reg

If reg Then
    'Regular return
    'If DontChangeFocus = 0 Then SetFocusWithoutErr txtField(fMinutesToComplete)
    'DontChangeFocus = DontChangeFocus + 1
    'chkEFile.Value = vbChecked
    'DontChangeFocus = DontChangeFocus - 1
Else
    'IPTE return
    txtField(fResultFederal).Text = ""
    txtField(fResultState).Text = ""
    txtField(fResultState2).Text = ""
    'DontChangeFocus = DontChangeFocus + 1
    'chkEFile.Value = vbChecked
    'DontChangeFocus = DontChangeFocus - 1
    txtField(fStateList).Text = ""

    txtField(fResultAGI).Text = ""

    'If DontChangeFocus = 0 Then SetFocusWithoutErr txtField(fMinutesToComplete)
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkIncPtnrTrustEstate_Click", Err
End Sub

'EHT=Standard
Private Sub chkIncPtnrTrustEstate_GotFocus()
On Error GoTo ERR_HANDLER

HilightControl Me, chkIncPtnrTrustEstate

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkIncPtnrTrustEstate_GotFocus", Err
End Sub

'EHT=Standard
Private Sub chkIncPtnrTrustEstate_LostFocus()
On Error GoTo ERR_HANDLER

ClearControlHilight Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "chkIncPtnrTrustEstate_LostFocus", Err
End Sub

'[Mark] Do we need to replace this with something new?
'Private Sub chkNoStateReturn_Click()
'If Not chkNoStateReturn.Enabled Then Exit Sub
'
'Dim b As Boolean
'b = (chkNoStateReturn.Value = vbUnchecked)
'
'If Not DontChangeFlags Then SetFutureFlagIndicator (Not b), NoStateReturn
'
'EnableTextbox txtField(fStateList), b
'EnableTextbox txtField(fResultState), b
'EnableTextbox txtField(fResultState2), b
'
'If b Then
'    If DontChangeFocus = 0 Then SetFocusWithoutErr txtField(fResultState)
'Else
'    txtField(fStateList).Text = ""
'    txtField(fResultState).Text = ""
'    txtField(fResultState2).Text = ""
'
'    If DontChangeFocus = 0 Then SetFocusWithoutErr txtField(fResultAGI)
'End If
'End Sub

'EHT=Standard
Private Sub lblFlagLastYear_Click(Index As Integer)
On Error GoTo ERR_HANDLER

SetLYFlagIndicator Not IsLYFlagIndicatorSet(0, Index), 0, Index

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lblFlagLastYear_Click", Err
End Sub

'EHT=Standard
Private Sub lblFlagThisYear_Click(Index As Integer)
On Error GoTo ERR_HANDLER

SetCYFlagIndicator Not IsCYFlagIndicatorSet(0, Index), 0, Index

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lblFlagThisYear_Click", Err
End Sub

'EHT=Standard
Private Sub optInType_Click(Index As Integer)
On Error GoTo ERR_HANDLER

Dim a%, c&

For a = 0 To optInType.UBound
    If a = Index Then
        Select Case a
        Case 0: c = &HC000&
        Case 1: c = &H88A5D5
        Case 2: c = &HECBA84
        Case 3: c = &H808080
        End Select
        optInType(a).BackColor = c
    Else
        optInType(a).BackColor = vbButtonFace
    End If
Next a

If Not optInType(Index).Enabled Then Exit Sub

If Not DontChangeFlags Then
    SetFutureFlagIndicator optInType(0).Value, HadAppointment
    SetFutureFlagIndicator optInType(1).Value, DroppedOff
    SetFutureFlagIndicator optInType(2).Value, MailedIn
    SetFutureFlagIndicator optInType(3).Value, NoNeedToFile
    SetFutureFlagIndicator Not optInType(3).Value, CompletedReturn
End If

'If chkIncPtnrTrustEstate is checked, then we don't want to mess with enabling/disabling
' anything, because it would interfere with what chkIncPtnrTrustEstate did
If chkIncPtnrTrustEstate.Value = vbUnchecked Then
    Dim b As Boolean
    b = Not optInType(3).Value

    EnableTextbox txtField(fMinutesToComplete), b

    EnableTextbox txtField(fCompletionDate), b
    EnableTextbox txtField(fPrepFee), b
    EnableTextbox txtField(fMoneyOwed), b

    EnableTextbox txtField(fResultFederal), b
    EnableTextbox txtField(fResultState), b
    EnableTextbox txtField(fResultState2), b
    chkEFile.Enabled = b
    EnableTextbox txtField(fStateList), b

    EnableTextbox txtField(fResultAGI), b

    chkIncPtnrTrustEstate.Enabled = b

    If b Then
        DontChangeFocus = DontChangeFocus + 1
        chkEFile.Value = vbChecked
        DontChangeFocus = DontChangeFocus - 1
    Else
        txtField(fMinutesToComplete).Text = ""

        txtField(fCompletionDate).Text = ""
        txtField(fPrepFee).Text = ""
        txtField(fMoneyOwed).Text = ""

        txtField(fResultFederal).Text = ""
        txtField(fResultState).Text = ""
        txtField(fResultState2).Text = ""
        DontChangeFocus = DontChangeFocus + 1
        chkEFile.Value = vbUnchecked
        DontChangeFocus = DontChangeFocus - 1
        txtField(fStateList).Text = ""

        txtField(fResultAGI).Text = ""
    End If
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "optInType_Click", Err
End Sub

'EHT=Standard
Private Sub optInType_GotFocus(Index As Integer)
On Error GoTo ERR_HANDLER

HilightControl Me, optInType(Index)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "optInType_GotFocus", Err
End Sub

'EHT=Standard
Private Sub optInType_LostFocus(Index As Integer)
On Error GoTo ERR_HANDLER

ClearControlHilight Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "optInType_LostFocus", Err
End Sub

'EHT=Standard
Private Sub txtField_Change(Index As Integer)
On Error GoTo ERR_HANDLER

Select Case Index
Case fResultFederal, fResultState, fResultState2, fResultAGI
    Dim n&
    With txtField(Index)
        FieldFromTextbox txtField(Index), n
        If n = NullLong Then
            .ForeColor = vbWindowText
        Else
            If n < 0 Then
                .ForeColor = &HC0&      'Red
            ElseIf n > 0 Then
                .ForeColor = &H8000&    'Green
            Else
                .ForeColor = vbWindowText
            End If
        End If
    End With
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_Change", Err
End Sub

'EHT=Standard
Private Sub txtField_GotFocus(Index As Integer)
On Error GoTo ERR_HANDLER

HilightControl Me, txtField(Index)
If (ShowFormMode = fPost) Then
    Select Case Index
    Case fMoneyOwed
        If txtField(fMoneyOwed).Text = "" Then
            If optInType(1).Value Or optInType(2).Value Or PreviouslyMarkedIncomplete Then
                txtField(fMoneyOwed).Text = txtField(fPrepFee).Text
            End If
        End If
    Case fCompletionDate
        If txtField(fCompletionDate).Text = "" Then
            FieldToTextbox txtField(fCompletionDate), Date
        End If
    End Select
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "txtField_GotFocus", Err
End Sub

'EHT=Standard
Private Sub txtField_LostFocus(Index As Integer)
On Error GoTo ERR_HANDLER

ClearControlHilight Me

Select Case Index
Case fPerson2Last
    'If Person2 has the same last name as Person1, then remove it
    If LCase$(txtField(fPerson2Last).Text) = LCase$(txtField(fPerson1Last)) Then
        txtField(fPerson2Last).Text = ""
    End If
Case fPerson1DOB, fPerson1DOD, fPerson2DOB, fPerson2DOD
    UpdateDOBandDODtext
End Select

LostFocusFormat txtField(Index)

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

'EHT=Standard
Sub UpdateDOBandDODtext()
On Error GoTo ERR_HANDLER

Dim a&, DOB&, DOD&
For a = 1 To 2
    If a = 1 Then
        FieldFromTextbox txtField(fPerson1DOB), DOB
        FieldFromTextbox txtField(fPerson1DOD), DOD
    Else
        FieldFromTextbox txtField(fPerson2DOB), DOB
        FieldFromTextbox txtField(fPerson2DOD), DOD
    End If
    If DOD <> NullLong Then
        If DOB = NullLong Then
            lblAge(a).Caption = "Died " & CalculateAge(DOD, Date) & "yr ago"
        Else
            If DOD >= DOB Then
                lblAge(a).Caption = "Died at age " & CalculateAge(DOB, DOD)
            Else
                lblAge(a).Caption = "ERR"
            End If
        End If
        lblAge(a).ForeColor = &HC0&     'Red
    ElseIf DOB <> NullLong Then
        If DOB <= Date Then
            lblAge(a).Caption = CalculateAge(DOB, Date) & "yr old today"
            lblAge(a).ForeColor = &H8000&   'Green
        Else
            lblAge(a).Caption = "ERR"
            lblAge(a).ForeColor = &HC0&     'Red
        End If
    Else
        lblAge(a).Caption = ""
    End If
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "UpdateDOBandDODtext", Err
End Sub

'EHT=Standard
Function IsCYFlagIndicatorSet(ActualFlag As ClientFlags, Optional ByVal IndIndex& = -1) As Boolean
On Error GoTo ERR_HANDLER

If IndIndex = -1 Then IndIndex = Log(ActualFlag) / Log(2)    'Convert flag back to a linear index
IsCYFlagIndicatorSet = (lblFlagThisYear(IndIndex).BackStyle = 1)

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "IsCYFlagIndicatorSet", Err
End Function

'EHT=Standard
Function IsLYFlagIndicatorSet(ActualFlag As ClientFlags, Optional ByVal IndIndex& = -1) As Boolean
On Error GoTo ERR_HANDLER

If IndIndex = -1 Then IndIndex = Log(ActualFlag) / Log(2)    'Convert flag back to a linear index
IsLYFlagIndicatorSet = (lblFlagLastYear(IndIndex).BackStyle = 1)

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "IsLYFlagIndicatorSet", Err
End Function

'EHT=Standard
Function IsFutureFlagIndicatorSet(ActualFlag As ClientFlags, Optional ByVal IndIndex& = -1) As Boolean
On Error GoTo ERR_HANDLER

If IndIndex = -1 Then IndIndex = Log(ActualFlag) / Log(2)    'Convert flag back to a linear index
IsFutureFlagIndicatorSet = (lblFlagFuture(IndIndex).BackStyle = 1)

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "IsFutureFlagIndicatorSet", Err
End Function

'EHT=Standard
Sub SetCYFlagIndicator(fset As Boolean, ActualFlag As ClientFlags, Optional ByVal IndIndex& = -1)
On Error GoTo ERR_HANDLER

If (ShowFormMode <> fNew) And (ShowFormMode <> fEdit) Then Err.Raise 1, , "Action only allowed in New or Edit modes!"
If IndIndex = -1 Then IndIndex = Log(ActualFlag) / Log(2)    'Convert flag back to a linear index
If fset Then
    lblFlagThisYear(IndIndex).BackStyle = 1
    lblFlagThisYear(IndIndex).BorderStyle = 1
Else
    lblFlagThisYear(IndIndex).BackStyle = 0
    lblFlagThisYear(IndIndex).BorderStyle = 0
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetCYFlagIndicator", Err
End Sub

'EHT=Standard
Sub SetLYFlagIndicator(fset As Boolean, ActualFlag As ClientFlags, Optional ByVal IndIndex& = -1)
On Error GoTo ERR_HANDLER

If (ShowFormMode <> fNew) And (ShowFormMode <> fEdit) Then Err.Raise 1, , "Action only allowed in New or Edit modes!"
If IndIndex = -1 Then IndIndex = Log(ActualFlag) / Log(2)    'Convert flag back to a linear index
If fset Then
    lblFlagLastYear(IndIndex).BackStyle = 1
    lblFlagLastYear(IndIndex).BorderStyle = 1
Else
    lblFlagLastYear(IndIndex).BackStyle = 0
    lblFlagLastYear(IndIndex).BorderStyle = 0
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetLYFlagIndicator", Err
End Sub

'EHT=Standard
Sub SetFutureFlagIndicator(fset As Boolean, ActualFlag As ClientFlags, Optional ByVal IndIndex& = -1)
On Error GoTo ERR_HANDLER

If (ShowFormMode <> fPost) Then Err.Raise 1, , "Action only allowed in Post mode!"
If IndIndex = -1 Then IndIndex = Log(ActualFlag) / Log(2)    'Convert flag back to a linear index
If fset Then
    lblFlagFuture(IndIndex).BackStyle = 1
    lblFlagFuture(IndIndex).BorderStyle = 1
Else
    lblFlagFuture(IndIndex).BackStyle = 0
    lblFlagFuture(IndIndex).BorderStyle = 0
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "SetFutureFlagIndicator", Err
End Sub

