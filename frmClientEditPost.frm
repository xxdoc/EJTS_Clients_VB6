VERSION 5.00
Begin VB.Form frmClientEditPost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit/Post Client"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14175
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
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctAppointmentHistory 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   2655
      TabIndex        =   102
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CheckBox chkField 
      Enabled         =   0   'False
      Height          =   375
      Index           =   712
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   6360
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.CheckBox chkField 
      Enabled         =   0   'False
      Height          =   375
      Index           =   705
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   3000
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   612
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   6360
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   605
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   3000
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.CheckBox chkField 
      Enabled         =   0   'False
      Height          =   375
      Index           =   702
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   1560
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   602
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   1560
      Value           =   2  'Grayed
      Width           =   735
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   507
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   5520
      Value           =   2  'Grayed
      Width           =   1815
   End
   Begin VB.ComboBox cboField 
      Height          =   360
      Index           =   600
      ItemData        =   "frmClientEditPost.frx":000C
      Left            =   11520
      List            =   "frmClientEditPost.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   87
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Index           =   700
      ItemData        =   "frmClientEditPost.frx":002B
      Left            =   12840
      List            =   "frmClientEditPost.frx":0038
      Style           =   2  'Dropdown List
      TabIndex        =   86
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Index           =   701
      ItemData        =   "frmClientEditPost.frx":004A
      Left            =   12840
      List            =   "frmClientEditPost.frx":005A
      Style           =   2  'Dropdown List
      TabIndex        =   85
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox cboField 
      Height          =   360
      Index           =   601
      ItemData        =   "frmClientEditPost.frx":0075
      Left            =   11520
      List            =   "frmClientEditPost.frx":0085
      Style           =   2  'Dropdown List
      TabIndex        =   83
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   505
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5520
      Value           =   2  'Grayed
      Width           =   1575
   End
   Begin VB.ComboBox cboField 
      Height          =   360
      Index           =   506
      ItemData        =   "frmClientEditPost.frx":00A0
      Left            =   3120
      List            =   "frmClientEditPost.frx":00B0
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   704
      Left            =   12840
      TabIndex        =   77
      Tag             =   "12"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   703
      Left            =   12840
      TabIndex        =   76
      Tag             =   "31"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   709
      Left            =   12840
      TabIndex        =   75
      Tag             =   "54"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   710
      Left            =   12840
      TabIndex        =   74
      Tag             =   "23"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   711
      Left            =   12840
      TabIndex        =   73
      Tag             =   "23"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   706
      Left            =   12840
      TabIndex        =   72
      Tag             =   "21"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   708
      Left            =   12840
      TabIndex        =   71
      Tag             =   "21"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtField 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
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
      Index           =   707
      Left            =   12840
      TabIndex        =   70
      Tag             =   "21"
      Top             =   3960
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
      Index           =   205
      Left            =   4080
      TabIndex        =   67
      Tag             =   "31"
      Top             =   3240
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
      Index           =   105
      Left            =   4080
      TabIndex        =   65
      Tag             =   "31"
      Top             =   1440
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
      Index           =   511
      Left            =   11280
      TabIndex        =   31
      Tag             =   "70"
      Top             =   7320
      Width           =   735
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
      Index           =   510
      Left            =   10320
      TabIndex        =   30
      Tag             =   "70"
      Top             =   7320
      Width           =   735
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
      Index           =   206
      Left            =   5640
      TabIndex        =   13
      Tag             =   "31"
      Top             =   3240
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
      Index           =   202
      Left            =   4560
      TabIndex        =   10
      Tag             =   "51"
      Top             =   2520
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
      Index           =   201
      Left            =   2760
      TabIndex        =   9
      Tag             =   "50"
      Top             =   2520
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
      Index           =   203
      Left            =   5640
      TabIndex        =   11
      Tag             =   "50"
      Top             =   2520
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
      Index           =   200
      Left            =   120
      TabIndex        =   8
      Tag             =   "50"
      Top             =   2520
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
      Left            =   5640
      TabIndex        =   6
      Tag             =   "31"
      Top             =   1440
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
      Index           =   102
      Left            =   4560
      TabIndex        =   3
      Tag             =   "51"
      Top             =   720
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
      Index           =   101
      Left            =   2760
      TabIndex        =   2
      Tag             =   "50"
      Top             =   720
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
      Index           =   204
      Left            =   120
      TabIndex        =   12
      Tag             =   "52"
      Top             =   3240
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
      Index           =   607
      Left            =   11520
      TabIndex        =   26
      Tag             =   "21"
      Top             =   3960
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
      Index           =   608
      Left            =   11520
      TabIndex        =   27
      Tag             =   "21"
      Top             =   4440
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
      Index           =   606
      Left            =   11520
      TabIndex        =   28
      Tag             =   "21"
      Top             =   3480
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
      Index           =   611
      Left            =   11520
      TabIndex        =   25
      Tag             =   "23"
      Top             =   5880
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
      Index           =   610
      Left            =   11520
      TabIndex        =   24
      Tag             =   "23"
      Top             =   5400
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
      Index           =   609
      Left            =   11520
      TabIndex        =   29
      Tag             =   "54"
      Top             =   4920
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
      Index           =   603
      Left            =   11520
      TabIndex        =   22
      Tag             =   "31"
      Text            =   "12/31/2015"
      Top             =   2040
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
      Index           =   604
      Left            =   11520
      TabIndex        =   23
      Tag             =   "12"
      Top             =   2520
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
      Index           =   504
      Left            =   120
      TabIndex        =   21
      Tag             =   "13"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtField 
      Height          =   615
      Index           =   509
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Tag             =   "50"
      Top             =   6240
      Width           =   6375
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
      Left            =   120
      TabIndex        =   5
      Tag             =   "52"
      Top             =   1440
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
      Index           =   503
      Left            =   6840
      TabIndex        =   18
      Tag             =   "51"
      Top             =   4800
      Width           =   2415
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
      Index           =   502
      Left            =   6240
      TabIndex        =   17
      Tag             =   "51"
      Top             =   4800
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
      Index           =   501
      Left            =   120
      TabIndex        =   16
      Tag             =   "51"
      Top             =   4800
      Width           =   6015
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
      Index           =   500
      Left            =   120
      TabIndex        =   15
      Tag             =   "51"
      Top             =   4320
      Width           =   9135
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
      Index           =   207
      Left            =   7200
      TabIndex        =   14
      Tag             =   "60"
      Top             =   3240
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
      Index           =   107
      Left            =   7200
      TabIndex        =   7
      Tag             =   "60"
      Top             =   1440
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
      Index           =   508
      Left            =   7200
      TabIndex        =   20
      Tag             =   "60"
      Top             =   5520
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
      Left            =   120
      TabIndex        =   1
      Tag             =   "50"
      Top             =   720
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
      Left            =   5640
      TabIndex        =   4
      Tag             =   "50"
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   7200
      TabIndex        =   33
      Top             =   7080
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
      Left            =   3720
      TabIndex        =   32
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label lblDODCalc 
      Alignment       =   2  'Center
      Caption         =   "Died at age 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   105
      Top             =   3630
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDODCalc 
      Alignment       =   2  'Center
      Caption         =   "Died at age 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   104
      Top             =   1830
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDOBCalc 
      Alignment       =   2  'Center
      Caption         =   "Died at age 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   103
      Top             =   3630
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Appointment history:"
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
      Index           =   25
      Left            =   120
      TabIndex        =   101
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label lbl 
      Caption         =   "Inc/Ptnr/Trust/Estate:"
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
      Left            =   5040
      TabIndex        =   100
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Reminder call:"
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
      Index           =   7
      Left            =   1200
      TabIndex        =   99
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lbl 
      Caption         =   "Mailing list mode:"
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
      Index           =   28
      Left            =   3120
      TabIndex        =   81
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "-"
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
      Index           =   26
      Left            =   11040
      TabIndex        =   79
      Top             =   7380
      Width           =   255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "2014"
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
      Index           =   24
      Left            =   12840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "2015"
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
      Index           =   14
      Left            =   11520
      TabIndex        =   78
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDOBCalc 
      Alignment       =   2  'Center
      Caption         =   "Died at age 100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   69
      Top             =   1830
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   4080
      TabIndex        =   68
      Top             =   3000
      Width           =   615
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
      Left            =   4080
      TabIndex        =   66
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "Years filed:"
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
      Left            =   10320
      TabIndex        =   63
      Top             =   7080
      Width           =   1695
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
      TabIndex        =   52
      Top             =   3720
      Width           =   9135
   End
   Begin VB.Label lbl 
      Caption         =   "Phone ('SP WORK'):"
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
      Left            =   7200
      TabIndex        =   49
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lbl 
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
      Left            =   5640
      TabIndex        =   48
      Top             =   3000
      Width           =   1455
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
      Left            =   4560
      TabIndex        =   45
      Top             =   2280
      Width           =   975
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
      Left            =   2760
      TabIndex        =   44
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "Email:"
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
      Left            =   120
      TabIndex        =   47
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label lbl 
      Caption         =   "Last (if different than above):"
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
      Left            =   5640
      TabIndex        =   43
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lbl 
      Caption         =   "First:"
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
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lbl 
      Caption         =   "Phone ('TP WORK'):"
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
      Left            =   7200
      TabIndex        =   40
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lbl 
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
      Left            =   5640
      TabIndex        =   41
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "Middle:"
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
      Left            =   4560
      TabIndex        =   35
      Top             =   480
      Width           =   975
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
      Left            =   2760
      TabIndex        =   36
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lbl 
      Caption         =   "First:"
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
      Left            =   120
      TabIndex        =   34
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lbl 
      Caption         =   "Last:"
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
      Left            =   5640
      TabIndex        =   37
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lbl 
      Caption         =   "Appt slots:"
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
      Left            =   120
      TabIndex        =   54
      Top             =   5280
      Width           =   975
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
      Left            =   2880
      TabIndex        =   53
      Top             =   6000
      Width           =   6375
   End
   Begin VB.Label lbl 
      Caption         =   "Email:"
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
      Left            =   120
      TabIndex        =   39
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label lbl 
      Caption         =   "Phone ('TP HOME'):"
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
      Left            =   7200
      TabIndex        =   51
      Top             =   5280
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
      Left            =   120
      TabIndex        =   50
      Top             =   4080
      Width           =   9135
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
      TabIndex        =   46
      Top             =   1920
      Width           =   9135
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
      TabIndex        =   38
      Top             =   120
      Width           =   9135
   End
   Begin VB.Line Line4 
      X1              =   624
      X2              =   624
      Y1              =   8
      Y2              =   456
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
      Left            =   13080
      TabIndex        =   64
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Filed extension:"
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
      Index           =   39
      Left            =   9360
      TabIndex        =   98
      Top             =   1620
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "E-Filed:"
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
      Index           =   38
      Left            =   9360
      TabIndex        =   97
      Top             =   3060
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Released before pmt:"
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
      Index           =   37
      Left            =   9360
      TabIndex        =   96
      Top             =   6420
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Inbox type:"
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
      Left            =   9360
      TabIndex        =   88
      Top             =   660
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Status:"
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
      Index           =   33
      Left            =   9360
      TabIndex        =   84
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "State(s):"
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
      Left            =   9360
      TabIndex        =   62
      Top             =   4980
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "State refund (or -due):"
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
      Left            =   9360
      TabIndex        =   61
      Top             =   4500
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Fed refund (or -due):"
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
      Left            =   9360
      TabIndex        =   60
      Top             =   4020
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
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
      Left            =   9360
      TabIndex        =   59
      Top             =   3540
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount still owed:"
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
      Left            =   9360
      TabIndex        =   58
      Top             =   5940
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Preparation fee:"
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
      Left            =   9360
      TabIndex        =   57
      Top             =   5460
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
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
      Left            =   9360
      TabIndex        =   56
      Top             =   2580
      Width           =   1935
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Date completed:"
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
      Left            =   9360
      TabIndex        =   55
      Top             =   2100
      Width           =   1935
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

Public Enum enumShowFormMode
    fPost
    fEdit
    fNew
End Enum

Public Enum enumClientTaxReturnFieldNum
    'Person #1
    fncPerson_First = 100
    fncPerson_Nickname
    fncPerson_Middle
    fncPerson_Last
    fncPerson_Email
    fncPerson_DateOfBirth
    fncPerson_DateOfDeath
    fncPerson_Phone
    
    'Person #2
    '(Same as above, but begins at 200)

    'Common
    fncMailingAddress_Street = 500
    fncMailingAddress_City
    fncMailingAddress_State
    fncMailingAddress_ZipCode
    fncNumApptSlots
    fncReminderCallAlways
    fncMailingListMode
    fncIPTE
    fncHomePhone
    fncNotes
    fncOldestYearFiled
    fncNewestYearFiled

    'TaxReturn #1
    fncInboxType = 600
    fncStatus
    fncFiledExtension
    fncCompletionDate
    fncMinutesToComplete
    fncEFiled
    fncResultAGI
    fncResultFederal
    fncResultState
    fncStateList
    fncFee
    fncFeeOwed
    fncReleasedBeforePayment
    
    'TaxReturn #2
    '(Same as above, but begins at 700)
End Enum

Private ShowFormMode As enumShowFormMode
Private PreviouslyMarkedIncomplete As Boolean
Private this As CClient
Private DataChanged As Boolean

'EHT=None
Private Sub Form_Load()
If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True
End Sub

'EHT=Cleanup2
Function Form_Show(vShowFormMode As enumShowFormMode, vClient As CClient, Optional vReadOnly As Boolean, Optional ByVal vOwnerForm As Form, Optional vNewClientInputString As String) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

'vShowFormMode              can be fPost, fEdit, or fNew
'vClient                    if fPost/fEdit, this is the CClient to open; if fNew, the new CClient will be set to this parameter (ByRef)
'vReadOnly                  if True, changes to vClient will not be allowed
'vOwnerForm                 only specify this if it is not frmMain or one of the tab 'forms'
'vNewClientInputString      only valid in fNew mode; initializes the new CClient with the specified data
'Return value               True if the CClient was changed in any way or if new client; False if Cancel button was used to close the form

'Copy some parameters to global for later access
ShowFormMode = vShowFormMode
If ShowFormMode = fNew Then
    Set this = New CClient
Else
    Set this = vClient
End If

'Basic form initialize
If ShowFormMode = fNew Then
    Me.Caption = "New Client"
Else
    Me.Caption = "Client #" & this.ID & " - " & Choose(ShowFormMode + 1, "Post", "Edit")
End If
btnSavePost.Caption = IIf(ShowFormMode = fPost, "&Post", "Save")
btnSavePost.Enabled = Not vReadOnly

'Populate the form with real data
this.PopulateForm Me
UpdateDOBandDODtext
DataChanged = False

'Set the tab order of controls
TabOrderSetting = IIf(ShowFormMode = fPost, "GLOBAL_TabOrder_ClientPost", "GLOBAL_TabOrder_ClientEdit")
SetControlTabOrder Me, DB_GetSetting(ActiveDBInstance, TabOrderSetting)

'Pause the main form's idle timer, so nothing changes in the background while we are on-screen
frmMain.IdlePauseTimeout

'Show ourselves, modal to the specified owner form
If vOwnerForm Is Nothing Then Set vOwnerForm = frmMain
Me.Show 1, vOwnerForm

'* * * * * * * * * CODE PAUSES AT THIS POINT UNTIL FORM IS CLOSED * * * * * * * * *

'Start the main form's idle timer again
frmMain.IdleSetAction

'Return values
Form_Show = DataChanged
If ShowFormMode = fNew Then vClient = this

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
    FieldFromTextbox txtField(fPerson1DOD), .Person1.dod

    FieldFromTextbox txtField(fPerson2First), .Person2.First
    FieldFromTextbox txtField(fPerson2Nickname), .Person2.Nickname
    FieldFromTextbox txtField(fPerson2Initial), .Person2.Initial
    FieldFromTextbox txtField(fPerson2Last), .Person2.Last
    FieldFromTextbox txtField(fPerson2Phone), .Person2.Phone
    FieldFromTextbox txtField(fPerson2Email), .Person2.Email
    FieldFromTextbox txtField(fPerson2DOB), .Person2.DOB
    FieldFromTextbox txtField(fPerson2DOD), .Person2.dod

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

If chkEFile.Value Then
    chkEFile.BackColor = &HC0FFC0
Else
    chkEFile.BackColor = vbButtonFace
End If

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

Dim a&, DOB&, dod&
For a = 1 To 2
    If a = 1 Then
        FieldFromTextbox txtField(fPerson1DOB), DOB
        FieldFromTextbox txtField(fPerson1DOD), dod
    Else
        FieldFromTextbox txtField(fPerson2DOB), DOB
        FieldFromTextbox txtField(fPerson2DOD), dod
    End If
    If dod <> NullLong Then
        If DOB = NullLong Then
            lblAge(a).Caption = "Died " & CalculateAge(dod, Date) & "yr ago"
        Else
            If dod >= DOB Then
                lblAge(a).Caption = "Died at age " & CalculateAge(DOB, dod)
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

