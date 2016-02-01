VERSION 5.00
Begin VB.Form frmClientEditPost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit/Post Client"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14055
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
   ScaleWidth      =   937
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctTaxReturn 
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   100
      Left            =   12660
      ScaleHeight     =   6255
      ScaleWidth      =   1335
      TabIndex        =   92
      Top             =   420
      Width           =   1335
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
         Left            =   60
         TabIndex        =   105
         Tag             =   "21"
         Top             =   3420
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
         Left            =   60
         TabIndex        =   104
         Tag             =   "21"
         Top             =   3900
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
         Left            =   60
         TabIndex        =   103
         Tag             =   "21"
         Top             =   2940
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
         Left            =   60
         TabIndex        =   102
         Tag             =   "23"
         Top             =   5340
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
         Left            =   60
         TabIndex        =   101
         Tag             =   "23"
         Top             =   4860
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
         Left            =   60
         TabIndex        =   100
         Tag             =   "54"
         Top             =   4380
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
         Left            =   60
         TabIndex        =   99
         Tag             =   "31"
         Top             =   1500
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
         Index           =   704
         Left            =   60
         TabIndex        =   98
         Tag             =   "12"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.ComboBox cboField 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   701
         ItemData        =   "frmClientEditPost.frx":000C
         Left            =   60
         List            =   "frmClientEditPost.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox cboField 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   700
         ItemData        =   "frmClientEditPost.frx":0037
         Left            =   60
         List            =   "frmClientEditPost.frx":0044
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   60
         Width           =   1215
      End
      Begin VB.CheckBox chkField 
         Enabled         =   0   'False
         Height          =   375
         Index           =   702
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1020
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Enabled         =   0   'False
         Height          =   375
         Index           =   705
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   2460
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Enabled         =   0   'False
         Height          =   375
         Index           =   712
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   5820
         Value           =   2  'Grayed
         Width           =   735
      End
   End
   Begin VB.PictureBox pctAppointmentHistory 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1575
      ScaleWidth      =   2655
      TabIndex        =   74
      Top             =   6240
      Width           =   2655
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   507
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5520
      Value           =   2  'Grayed
      Width           =   1815
   End
   Begin VB.CheckBox chkField 
      Height          =   375
      Index           =   505
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   5520
      Value           =   2  'Grayed
      Width           =   1575
   End
   Begin VB.ComboBox cboField 
      Height          =   360
      Index           =   506
      ItemData        =   "frmClientEditPost.frx":0056
      Left            =   3120
      List            =   "frmClientEditPost.frx":0066
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   5520
      Width           =   1575
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
      TabIndex        =   58
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
      TabIndex        =   56
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   12
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
      MaxLength       =   1
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   10
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      MaxLength       =   1
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   11
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
      Index           =   504
      Left            =   120
      TabIndex        =   20
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
      TabIndex        =   18
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
      TabIndex        =   4
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   6
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
      TabIndex        =   19
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
      TabIndex        =   0
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
      TabIndex        =   3
      Tag             =   "50"
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   7200
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   6960
      Width           =   3135
   End
   Begin VB.PictureBox pctTaxReturn 
      BorderStyle     =   0  'None
      Height          =   6255
      Index           =   0
      Left            =   11340
      ScaleHeight     =   6255
      ScaleWidth      =   1335
      TabIndex        =   78
      Top             =   420
      Width           =   1335
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
         Left            =   60
         TabIndex        =   91
         Tag             =   "12"
         Top             =   1980
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
         Left            =   60
         TabIndex        =   90
         Tag             =   "31"
         Top             =   1500
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
         Left            =   60
         TabIndex        =   89
         Tag             =   "54"
         Top             =   4380
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
         Left            =   60
         TabIndex        =   88
         Tag             =   "23"
         Top             =   4860
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
         Left            =   60
         TabIndex        =   87
         Tag             =   "23"
         Top             =   5340
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
         Left            =   60
         TabIndex        =   86
         Tag             =   "21"
         Top             =   2940
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
         Left            =   60
         TabIndex        =   85
         Tag             =   "21"
         Top             =   3900
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
         Index           =   607
         Left            =   60
         TabIndex        =   84
         Tag             =   "21"
         Top             =   3420
         Width           =   1215
      End
      Begin VB.ComboBox cboField 
         Height          =   360
         Index           =   601
         ItemData        =   "frmClientEditPost.frx":0088
         Left            =   60
         List            =   "frmClientEditPost.frx":0098
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   540
         Width           =   1215
      End
      Begin VB.ComboBox cboField 
         Height          =   360
         Index           =   600
         ItemData        =   "frmClientEditPost.frx":00B3
         Left            =   60
         List            =   "frmClientEditPost.frx":00C0
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   60
         Width           =   1215
      End
      Begin VB.CheckBox chkField 
         Height          =   375
         Index           =   602
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1020
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   375
         Index           =   605
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2460
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   375
         Index           =   612
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   5820
         Value           =   2  'Grayed
         Width           =   735
      End
   End
   Begin VB.Label lblSwitchPersons 
      AutoSize        =   -1  'True
      Caption         =   "ст"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      TabIndex        =   110
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label lblNoTaxReturn 
      Caption         =   "There is no tax return entered for this year."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Index           =   100
      Left            =   12720
      TabIndex        =   109
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblNoTaxReturn 
      Caption         =   "There is no tax return entered for this year."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Index           =   0
      Left            =   11400
      TabIndex        =   108
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblTaxYear 
      Alignment       =   2  'Center
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   100
      Left            =   12720
      TabIndex        =   107
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTaxYear 
      Alignment       =   2  'Center
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   11400
      TabIndex        =   106
      Top             =   120
      Width           =   1215
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   63
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
      TabIndex        =   61
      Top             =   7380
      Width           =   255
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
      TabIndex        =   60
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
      TabIndex        =   59
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
      TabIndex        =   57
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
      TabIndex        =   54
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
      TabIndex        =   43
      Top             =   3720
      Width           =   2535
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
      TabIndex        =   40
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
      TabIndex        =   39
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   38
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   31
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
      TabIndex        =   32
      Top             =   1200
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
      Index           =   1
      Left            =   4560
      TabIndex        =   26
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
      TabIndex        =   27
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
      TabIndex        =   25
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
      TabIndex        =   28
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
      TabIndex        =   45
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
      TabIndex        =   44
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
      TabIndex        =   30
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
      TabIndex        =   42
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
      TabIndex        =   41
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
      TabIndex        =   37
      Top             =   1920
      Width           =   2535
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
      TabIndex        =   29
      Top             =   120
      Width           =   2535
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
      Left            =   12960
      TabIndex        =   55
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
      Left            =   9390
      TabIndex        =   70
      Top             =   1500
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   69
      Top             =   2940
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   68
      Top             =   6300
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   66
      Top             =   540
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   65
      Top             =   1020
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   53
      Top             =   4860
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   52
      Top             =   4380
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   51
      Top             =   3900
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   50
      Top             =   3420
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   49
      Top             =   5820
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   48
      Top             =   5340
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   47
      Top             =   2460
      Width           =   1905
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
      Left            =   9390
      TabIndex        =   46
      Top             =   1980
      Width           =   1905
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
Private This As CClient
Private DataChanged As Boolean





'#################################################################################
'Load / Show / Save / Unload
'#################################################################################

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
    Set This = New CClient
Else
    Set This = vClient
End If

'Basic form initialize
If ShowFormMode = fNew Then
    Me.Caption = "New Client"
Else
    Me.Caption = "Client #" & This.ID & " - " & Choose(ShowFormMode + 1, "Post", "Edit")
End If
btnSavePost.Caption = IIf(ShowFormMode = fPost, "&Post", "Save")
btnSavePost.Enabled = Not vReadOnly

'Populate the form with real data
If Not This.PopulateToForm(Me) Then
    'An error occured, and the user was already notified, so just quit
    HASERROR = True: GoTo CLEANUP
End If
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
If ShowFormMode = fNew Then vClient = This

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Function

'EHT=Standard
Private Sub btnSavePost_Click()
On Error GoTo ERR_HANDLER

If Not btnSavePost.Enabled Then Exit Sub

If This.PopulateFromForm(Me) Then
    If ShowFormMode = fNew Then
        'Add it to the database
    End If
    frmMain.SetChangedFlagAndIndication
    DataChanged = True
    Unload Me
End If

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnSavePost_Click", Err
End Sub

'EHT=Standard
Private Sub btnCancel_Click()
On Error GoTo ERR_HANDLER

If Not btnCancel.Enabled Then Exit Sub

Unload Me

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "btnCancel_Click", Err
End Sub





'#################################################################################
'Tabbing between controls
'#################################################################################

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
Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ERR_HANDLER

Select Case KeyCode
Case vbKeyReturn
    If Shift = vbCtrlMask Then
        SetFocusWithoutErr btnSavePost
        btnSavePost_Click
    Else
        TabToNextControl Me, True, (Shift = vbShiftMask)
    End If
Case 65     'A
    If Shift = vbCtrlMask Then
        If TypeName(Me.ActiveControl) = "TextBox" Then
            'Select contents
            With Me.ActiveControl
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End If
End Select

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_KeyDown", Err
End Sub

'EHT=ResumeNext
Public Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

Select Case KeyAscii
Case vbKeyReturn
    KeyAscii = 0    'Stop the beep
End Select
End Sub





'#################################################################################
'Field handling: GotFocus/LostFocus/Click
'#################################################################################

'EHT=ResumeNext
Private Sub txtField_GotFocus(Index As Integer)
On Error Resume Next

HilightControl Me, txtField(Index)

If ShowFormMode = fPost Then
    Select Case Index
    Case fncFeeOwed
        If txtField(fncFeeOwed).Text = "" Then
            Dim li As Integer
            li = cboField(fncInboxType).ListIndex
            If li = itDroppedOff Or li = itMailedIn Then
                txtField(fncFeeOwed).Text = txtField(fncFee).Text
            End If
        End If
    Case fncCompletionDate
        If txtField(fncCompletionDate).Text = "" Then
            FieldToTextbox txtField(fncCompletionDate), Date
        End If
    End Select
End If
End Sub

'EHT=ResumeNext
Private Sub txtField_LostFocus(Index As Integer)
On Error Resume Next

Dim v As Variant, n$(1)

ClearControlHilight Me
If Not ValidateTextbox(txtField(Index), v) Then Exit Sub

Select Case Index
Case fncPerson_Last, fncPerson_Last + frmClientEditPost_PersonOffset
    n$(0) = LCase$(txtField(fncPerson_Last).Text)
    n$(1) = LCase$(txtField(fncPerson_Last + frmClientEditPost_PersonOffset).Text)
    'If same last names, second one should be greyed out
    If Len(n$(1)) > 0 And (n$(1) = n$(0)) Then
        'Same = Grey
        txtField(fncPerson_Last + frmClientEditPost_PersonOffset).ForeColor = &HC0C0C0
    Else
        'Different = Black
        txtField(fncPerson_Last + frmClientEditPost_PersonOffset).ForeColor = vbWindowText
    End If

Case fncPerson_DateOfBirth, fncPerson_DateOfDeath, fncPerson_DateOfBirth + frmClientEditPost_PersonOffset, fncPerson_DateOfDeath + frmClientEditPost_PersonOffset
    UpdateDOBandDODtext

Case fncResultAGI, fncResultFederal, fncResultState
    With txtField(Index)
        If v = NullLong Then
            'Blank = Black
            .ForeColor = vbWindowText
        Else
            If v < 0 Then
                'Negative = Red
                .ForeColor = &HC0&
            ElseIf v > 0 Then
                'Positive = Green
                .ForeColor = &H8000&    'Green
            Else
                'Zero = Black
                .ForeColor = vbWindowText
            End If
        End If
    End With
End Select
End Sub

'EHT=ResumeNext
Private Sub chkField_Click(Index As Integer)
On Error Resume Next

ValidateCheckbox chkField(Index), False

With chkField(Index)
    If .style = vbButtonGraphical Then
        If .Value = vbChecked Then
            .Caption = "yes"
            .FontBold = True
        ElseIf .Value = vbUnchecked Then
            .Caption = "no"
            .FontBold = False
        ElseIf .Value = vbGrayed Then
            .Caption = ""
        End If
    End If
End With
End Sub

'EHT=ResumeNext
Private Sub chkField_GotFocus(Index As Integer)
On Error Resume Next

HilightControl Me, chkField(Index)
End Sub

'EHT=ResumeNext
Private Sub chkField_LostFocus(Index As Integer)
On Error Resume Next

ClearControlHilight Me
End Sub

'EHT=ResumeNext
Private Sub cboField_Click(Index As Integer)
On Error Resume Next

ValidateCombobox cboField(Index), 0
End Sub

'EHT=ResumeNext
Private Sub cboField_GotFocus(Index As Integer)
On Error Resume Next

HilightControl Me, cboField(Index)
End Sub

'EHT=ResumeNext
Private Sub cboField_LostFocus(Index As Integer)
On Error Resume Next

ClearControlHilight Me
End Sub

'EHT=ResumeNext
Private Sub btnSavePost_GotFocus()
On Error Resume Next

HilightControl Me, btnSavePost
End Sub

'EHT=ResumeNext
Private Sub btnSavePost_LostFocus()
On Error Resume Next

ClearControlHilight Me
End Sub

'EHT=ResumeNext
Private Sub btnCancel_GotFocus()
On Error Resume Next

HilightControl Me, btnCancel
End Sub

'EHT=ResumeNext
Private Sub btnCancel_LostFocus()
On Error Resume Next

ClearControlHilight Me
End Sub





'#################################################################################
'Switching of person #1 and person #2
'#################################################################################

'EHT=Standard
Private Sub lblSwitchPersons_Click()
On Error GoTo ERR_HANDLER

SwitchTextboxValues fncPerson_First, fncPerson_First + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_Nickname, fncPerson_Nickname + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_Middle, fncPerson_Middle + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_Last, fncPerson_Last + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_Email, fncPerson_Email + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_DateOfBirth, fncPerson_DateOfBirth + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_DateOfDeath, fncPerson_DateOfDeath + frmClientEditPost_PersonOffset
SwitchTextboxValues fncPerson_Phone, fncPerson_Phone + frmClientEditPost_PersonOffset

UpdateDOBandDODtext

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lblSwitchPersons_Click", Err
End Sub

'EHT=None
Private Sub SwitchTextboxValues(txt1 As Integer, txt2 As Integer)
Dim v$
v$ = txtField(txt1).Text
txtField(txt1).Text = txtField(txt2).Text
txtField(txt2).Text = v$
End Sub





'#################################################################################
'Calculations for DOB and DOD
'#################################################################################

'EHT=Standard
Sub UpdateDOBandDODtext()
On Error GoTo ERR_HANDLER

Dim a&, DOB&, DOD&
For a = 0 To 1
    FieldFromTextbox txtField(fncPerson_DateOfBirth + (a * frmClientEditPost_PersonOffset)), DOB
    FieldFromTextbox txtField(fncPerson_DateOfDeath + (a * frmClientEditPost_PersonOffset)), DOD
    If DOD <> NullLong Then
        If DOB = NullLong Then
            lblDODCalc(a).Caption = "Died " & CalculateAge(DOD, Date) & "yr ago"
        Else
            If DOD >= DOB Then
                lblDODCalc(a).Caption = "Died at age " & CalculateAge(DOB, DOD)
            Else
                lblDODCalc(a).Caption = "ERR"
            End If
        End If
        lblDOBCalc(a).Visible = False
        lblDODCalc(a).Visible = True
    ElseIf DOB <> NullLong Then
        If DOB <= Date Then
            lblDOBCalc(a).Caption = CalculateAge(DOB, Date) & "yr old today"
        Else
            lblDOBCalc(a).Caption = "ERR"
        End If
        lblDOBCalc(a).Visible = True
        lblDODCalc(a).Visible = False
    Else
        lblDOBCalc(a).Visible = False
        lblDODCalc(a).Visible = False
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
