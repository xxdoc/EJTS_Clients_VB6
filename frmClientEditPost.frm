VERSION 5.00
Begin VB.Form frmClientEditPost 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit/Post Client"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   633
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkgroupField 
      Height          =   375
      Index           =   38
      Left            =   9060
      TabIndex        =   161
      Top             =   1440
      Value           =   2  'Grayed
      Width           =   255
   End
   Begin VB.CheckBox chkgroupField 
      Height          =   375
      Index           =   39
      Left            =   9060
      TabIndex        =   160
      Top             =   3600
      Value           =   2  'Grayed
      Width           =   255
   End
   Begin VB.CheckBox chkgroupField 
      Height          =   375
      Index           =   37
      Left            =   9060
      TabIndex        =   159
      Top             =   5040
      Value           =   2  'Grayed
      Width           =   255
   End
   Begin VB.CheckBox chkField 
      Height          =   255
      Index           =   30
      Left            =   8400
      TabIndex        =   156
      Tag             =   "YES|no"
      Top             =   120
      Value           =   2  'Grayed
      Width           =   855
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   46
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   155
      Tag             =   "50"
      Top             =   7560
      Width           =   8895
   End
   Begin VB.PictureBox pctTaxReturn 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   2
      Left            =   14520
      ScaleHeight     =   521
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   129
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CheckBox chkField 
         Height          =   255
         Index           =   97
         Left            =   360
         TabIndex        =   150
         Top             =   4800
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   255
         Index           =   102
         Left            =   360
         TabIndex        =   149
         Top             =   7080
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   255
         Index           =   103
         Left            =   360
         TabIndex        =   148
         Top             =   7440
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   94
         Left            =   120
         TabIndex        =   147
         Tag             =   "12"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   93
         Left            =   120
         TabIndex        =   146
         Tag             =   "31"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Left            =   120
         TabIndex        =   145
         Tag             =   "54"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   95
         Left            =   120
         TabIndex        =   144
         Tag             =   "23"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   96
         Left            =   120
         TabIndex        =   143
         Tag             =   "23"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   98
         Left            =   120
         TabIndex        =   142
         Tag             =   "21"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         TabIndex        =   141
         Tag             =   "21"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   99
         Left            =   120
         TabIndex        =   140
         Tag             =   "21"
         Top             =   5640
         Width           =   1215
      End
      Begin VB.PictureBox chsField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   885
         Index           =   90
         Left            =   120
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   136
         Top             =   120
         Width           =   1215
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Appointment"
            Height          =   285
            Index           =   18
            Left            =   0
            TabIndex        =   139
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Dropped off"
            Height          =   285
            Index           =   19
            Left            =   0
            TabIndex        =   138
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Mailed in"
            Height          =   285
            Index           =   20
            Left            =   0
            TabIndex        =   137
            Top             =   570
            Width           =   1185
         End
      End
      Begin VB.PictureBox chsField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1170
         Index           =   91
         Left            =   120
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   131
         Top             =   1080
         Width           =   1215
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "NNTF"
            Height          =   285
            Index           =   24
            Left            =   0
            TabIndex        =   135
            Top             =   855
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Complete"
            Height          =   285
            Index           =   23
            Left            =   0
            TabIndex        =   134
            Top             =   570
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Incomplete"
            Height          =   285
            Index           =   22
            Left            =   0
            TabIndex        =   133
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Not started"
            Height          =   285
            Index           =   21
            Left            =   0
            TabIndex        =   132
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   92
         Left            =   120
         TabIndex        =   130
         Tag             =   "31"
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.PictureBox chsField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1170
      Index           =   44
      Left            =   6960
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   135
      TabIndex        =   112
      Top             =   5760
      Width           =   2055
      Begin VB.Label lblChooserChoice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC98C&
         Caption         =   "Hard-copy organizer"
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   116
         Top             =   855
         Width           =   2025
      End
      Begin VB.Label lblChooserChoice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC98C&
         Caption         =   "Email organizer"
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   115
         Top             =   570
         Width           =   2025
      End
      Begin VB.Label lblChooserChoice 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC98C&
         Caption         =   "No organizer"
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   114
         Top             =   285
         Width           =   2025
      End
      Begin VB.Label lblChooserChoice 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Auto"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   113
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.PictureBox pctTaxReturn 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   1
      Left            =   12960
      ScaleHeight     =   521
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   87
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   72
         Left            =   120
         TabIndex        =   128
         Tag             =   "31"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.PictureBox chsField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1170
         Index           =   71
         Left            =   120
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   122
         Top             =   1080
         Width           =   1215
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Not started"
            Height          =   285
            Index           =   14
            Left            =   0
            TabIndex        =   126
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Incomplete"
            Height          =   285
            Index           =   15
            Left            =   0
            TabIndex        =   125
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Complete"
            Height          =   285
            Index           =   16
            Left            =   0
            TabIndex        =   124
            Top             =   570
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "NNTF"
            Height          =   285
            Index           =   17
            Left            =   0
            TabIndex        =   123
            Top             =   855
            Width           =   1185
         End
      End
      Begin VB.PictureBox chsField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   885
         Index           =   70
         Left            =   120
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   108
         Top             =   120
         Width           =   1215
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Mailed in"
            Height          =   285
            Index           =   13
            Left            =   0
            TabIndex        =   111
            Top             =   570
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Dropped off"
            Height          =   285
            Index           =   12
            Left            =   0
            TabIndex        =   110
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00A5A5A5&
            Caption         =   "Appointment"
            Height          =   285
            Index           =   11
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   79
         Left            =   120
         TabIndex        =   98
         Tag             =   "21"
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   80
         Left            =   120
         TabIndex        =   97
         Tag             =   "21"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   78
         Left            =   120
         TabIndex        =   96
         Tag             =   "21"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   76
         Left            =   120
         TabIndex        =   95
         Tag             =   "23"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   75
         Left            =   120
         TabIndex        =   94
         Tag             =   "23"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   81
         Left            =   120
         TabIndex        =   93
         Tag             =   "54"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   73
         Left            =   120
         TabIndex        =   92
         Tag             =   "31"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
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
         Index           =   74
         Left            =   120
         TabIndex        =   91
         Tag             =   "12"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox chkField 
         Height          =   255
         Index           =   83
         Left            =   360
         TabIndex        =   90
         Top             =   7440
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   255
         Index           =   82
         Left            =   360
         TabIndex        =   89
         Top             =   7080
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   255
         Index           =   77
         Left            =   360
         TabIndex        =   88
         Top             =   4800
         Value           =   2  'Grayed
         Width           =   735
      End
   End
   Begin VB.CheckBox chkField 
      Height          =   255
      Index           =   31
      Left            =   5400
      TabIndex        =   66
      Tag             =   "YES|no"
      Top             =   120
      Value           =   2  'Grayed
      Width           =   855
   End
   Begin VB.CheckBox chkField 
      Height          =   360
      Index           =   41
      Left            =   1440
      TabIndex        =   63
      Top             =   6240
      Value           =   2  'Grayed
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   26
      Left            =   3840
      TabIndex        =   58
      Tag             =   "31"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   3840
      TabIndex        =   56
      Tag             =   "31"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   43
      Left            =   5520
      TabIndex        =   22
      Tag             =   "70"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   42
      Left            =   4560
      TabIndex        =   21
      Tag             =   "70"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   27
      Left            =   5400
      TabIndex        =   12
      Tag             =   "31"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   22
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   9
      Tag             =   "51"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   21
      Left            =   2760
      TabIndex        =   8
      Tag             =   "50"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   23
      Left            =   5640
      TabIndex        =   10
      Tag             =   "50"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   20
      Left            =   120
      TabIndex        =   7
      Tag             =   "50"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   5400
      TabIndex        =   5
      Tag             =   "31"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   2
      Tag             =   "51"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Tag             =   "50"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   25
      Left            =   120
      TabIndex        =   11
      Tag             =   "52"
      Top             =   3600
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   40
      Left            =   120
      TabIndex        =   20
      Tag             =   "13"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   45
      Left            =   120
      TabIndex        =   18
      Tag             =   "50"
      Top             =   7080
      Width           =   8895
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Tag             =   "52"
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   35
      Left            =   5520
      TabIndex        =   17
      Tag             =   "51"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   34
      Left            =   4920
      TabIndex        =   16
      Tag             =   "51"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   33
      Left            =   120
      TabIndex        =   15
      Tag             =   "51"
      Top             =   5520
      Width           =   4695
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   32
      Left            =   120
      TabIndex        =   14
      Tag             =   "51"
      Top             =   5040
      Width           =   6735
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   24
      Left            =   6960
      TabIndex        =   13
      Tag             =   "60"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   6960
      TabIndex        =   6
      Tag             =   "60"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   36
      Left            =   6960
      TabIndex        =   19
      Tag             =   "60"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "50"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtField 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   5640
      TabIndex        =   3
      Tag             =   "50"
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5640
      TabIndex        =   24
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CommandButton btnSave 
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
      Left            =   2160
      TabIndex        =   23
      Top             =   8520
      Width           =   3135
   End
   Begin VB.PictureBox pctTaxReturn 
      BorderStyle     =   0  'None
      Height          =   7815
      Index           =   0
      Left            =   11400
      ScaleHeight     =   521
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   75
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   52
         Left            =   120
         TabIndex        =   158
         Tag             =   "31"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.PictureBox chsField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1170
         Index           =   51
         Left            =   120
         ScaleHeight     =   76
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   117
         Top             =   1080
         Width           =   1215
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H000080FF&
            Caption         =   "NNTF"
            Height          =   285
            Index           =   10
            Left            =   0
            TabIndex        =   121
            Top             =   855
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H0000C000&
            Caption         =   "Complete"
            Height          =   285
            Index           =   9
            Left            =   0
            TabIndex        =   120
            Top             =   570
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H0007C9E4&
            Caption         =   "Incomplete"
            Height          =   285
            Index           =   8
            Left            =   0
            TabIndex        =   119
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC98C&
            Caption         =   "Not started"
            Height          =   285
            Index           =   7
            Left            =   0
            TabIndex        =   118
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.PictureBox chsField 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   885
         Index           =   50
         Left            =   120
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   104
         Top             =   120
         Width           =   1215
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC98C&
            Caption         =   "Appointment"
            Height          =   285
            Index           =   4
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC98C&
            Caption         =   "Dropped off"
            Height          =   285
            Index           =   5
            Left            =   0
            TabIndex        =   106
            Top             =   285
            Width           =   1185
         End
         Begin VB.Label lblChooserChoice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC98C&
            Caption         =   "Mailed in"
            Height          =   285
            Index           =   6
            Left            =   0
            TabIndex        =   105
            Top             =   570
            Width           =   1185
         End
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   54
         Left            =   120
         TabIndex        =   86
         Tag             =   "12"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   53
         Left            =   120
         TabIndex        =   85
         Tag             =   "31"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   61
         Left            =   120
         TabIndex        =   84
         Tag             =   "54"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   55
         Left            =   120
         TabIndex        =   83
         Tag             =   "23"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   56
         Left            =   120
         TabIndex        =   82
         Tag             =   "23"
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   58
         Left            =   120
         TabIndex        =   81
         Tag             =   "21"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   60
         Left            =   120
         TabIndex        =   80
         Tag             =   "21"
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   59
         Left            =   120
         TabIndex        =   79
         Tag             =   "21"
         Top             =   5640
         Width           =   1215
      End
      Begin VB.CheckBox chkField 
         Height          =   240
         Index           =   63
         Left            =   360
         TabIndex        =   78
         Top             =   7440
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   240
         Index           =   62
         Left            =   360
         TabIndex        =   77
         Top             =   7080
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkField 
         Height          =   240
         Index           =   57
         Left            =   360
         TabIndex        =   76
         Top             =   4800
         Value           =   2  'Grayed
         Width           =   735
      End
   End
   Begin VB.Label lblSwapPersons 
      AutoSize        =   -1  'True
      Caption         =   "swap"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7815
      TabIndex        =   165
      ToolTipText     =   "Swap data between Taxpayer and Spouse"
      Top             =   2130
      Width           =   705
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "BEST #"
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
      Index           =   37
      Left            =   8640
      TabIndex        =   164
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "BEST #"
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
      Index           =   39
      Left            =   8640
      TabIndex        =   163
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "BEST #"
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
      Index           =   38
      Left            =   8640
      TabIndex        =   162
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Unused client record"
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
      Left            =   6480
      TabIndex        =   157
      Top             =   135
      Width           =   1815
   End
   Begin VB.Image imgTRDivider 
      Height          =   7785
      Left            =   12855
      Picture         =   "frmClientEditPost.frx":000C
      Top             =   480
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblShiftTRView 
      Alignment       =   2  'Center
      Caption         =   "»"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   14520
      TabIndex        =   154
      ToolTipText     =   "Shift the view to see older tax returns"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblShiftTRView 
      Alignment       =   2  'Center
      Caption         =   "«"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   12960
      TabIndex        =   153
      ToolTipText     =   "Shift the view to see more recent tax returns"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
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
      Index           =   2
      Left            =   14520
      TabIndex        =   152
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblNoTaxReturn 
      Enabled         =   0   'False
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
      Left            =   14520
      TabIndex        =   151
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Appointment date:"
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
      Index           =   52
      Left            =   9480
      TabIndex        =   127
      Top             =   2925
      Width           =   1815
   End
   Begin VB.Label lblSwapPersons 
      AutoSize        =   -1  'True
      Caption         =   "αβ"
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
      Index           =   0
      Left            =   8520
      TabIndex        =   103
      Top             =   2160
      Width           =   510
   End
   Begin VB.Label lblNoTaxReturn 
      Enabled         =   0   'False
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
      Left            =   12960
      TabIndex        =   102
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblNoTaxReturn 
      Enabled         =   0   'False
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
      Left            =   11400
      TabIndex        =   101
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
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
      Index           =   1
      Left            =   12960
      TabIndex        =   100
      Top             =   120
      Width           =   1455
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
      TabIndex        =   99
      Top             =   120
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
      Index           =   1
      Left            =   5400
      TabIndex        =   74
      Top             =   3990
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDODCalc 
      Alignment       =   2  'Center
      Caption         =   "Died 100yr ago"
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
      Left            =   5400
      TabIndex        =   73
      Top             =   1830
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblDOBCalc 
      Alignment       =   2  'Center
      Caption         =   "ERR"
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
      Left            =   3840
      TabIndex        =   72
      Top             =   3990
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
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
      Index           =   31
      Left            =   3480
      TabIndex        =   71
      Top             =   135
      Width           =   1815
   End
   Begin VB.Label lblField 
      Caption         =   "Always remind?"
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
      Index           =   41
      Left            =   1440
      TabIndex        =   70
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label lblField 
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
      Index           =   44
      Left            =   6960
      TabIndex        =   62
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label lblYearRang 
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
      Left            =   5280
      TabIndex        =   61
      Top             =   6300
      Width           =   255
   End
   Begin VB.Label lblDOBCalc 
      Alignment       =   2  'Center
      Caption         =   "100yr old today"
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
      Left            =   3840
      TabIndex        =   60
      Top             =   1830
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblField 
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
      Index           =   26
      Left            =   3840
      TabIndex        =   59
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblField 
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
      Index           =   6
      Left            =   3840
      TabIndex        =   57
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblField 
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
      Index           =   42
      Left            =   4560
      TabIndex        =   54
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label lblTitle 
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
      Index           =   2
      Left            =   120
      TabIndex        =   43
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label lblField 
      Caption         =   "Cell phone:"
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
      Index           =   24
      Left            =   6960
      TabIndex        =   40
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblField 
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
      Index           =   27
      Left            =   5400
      TabIndex        =   39
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblField 
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
      Index           =   22
      Left            =   4560
      TabIndex        =   36
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblField 
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
      Index           =   21
      Left            =   2760
      TabIndex        =   35
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblField 
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
      Index           =   25
      Left            =   120
      TabIndex        =   38
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblField 
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
      Index           =   23
      Left            =   5640
      TabIndex        =   34
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label lblField 
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
      Index           =   20
      Left            =   120
      TabIndex        =   33
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblField 
      Caption         =   "Cell phone:"
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
      Left            =   6960
      TabIndex        =   31
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblField 
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
      Index           =   7
      Left            =   5400
      TabIndex        =   32
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblField 
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
      Index           =   2
      Left            =   4560
      TabIndex        =   26
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblField 
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
      Index           =   1
      Left            =   2760
      TabIndex        =   27
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblField 
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
   Begin VB.Label lblField 
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
      Width           =   3375
   End
   Begin VB.Label lblField 
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
      Index           =   40
      Left            =   120
      TabIndex        =   45
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblField 
      Caption         =   "Notes (only the top field will show on search tab):"
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
      Index           =   45
      Left            =   120
      TabIndex        =   44
      Top             =   6840
      Width           =   6975
   End
   Begin VB.Label lblField 
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
      Width           =   3615
   End
   Begin VB.Label lblField 
      Caption         =   "Home phone:"
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
      Left            =   6960
      TabIndex        =   42
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblField 
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
      Index           =   32
      Left            =   120
      TabIndex        =   41
      Top             =   4800
      Width           =   6735
   End
   Begin VB.Label lblTitle 
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
      Index           =   1
      Left            =   120
      TabIndex        =   37
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblTitle 
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
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   2535
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
      Left            =   15000
      TabIndex        =   55
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label lblField 
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
      Index           =   63
      Left            =   9480
      TabIndex        =   69
      Top             =   7935
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   62
      Left            =   9480
      TabIndex        =   68
      Top             =   7575
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   57
      Left            =   9480
      TabIndex        =   67
      Top             =   5295
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   50
      Left            =   9480
      TabIndex        =   65
      Top             =   630
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   51
      Left            =   9480
      TabIndex        =   64
      Top             =   1590
      Width           =   1815
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "State return(s):"
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
      Index           =   61
      Left            =   9480
      TabIndex        =   53
      Top             =   7125
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   60
      Left            =   9480
      TabIndex        =   52
      Top             =   6645
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   59
      Left            =   9480
      TabIndex        =   51
      Top             =   6165
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   58
      Left            =   9480
      TabIndex        =   50
      Top             =   5685
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   56
      Left            =   9480
      TabIndex        =   49
      Top             =   4845
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   55
      Left            =   9480
      TabIndex        =   48
      Top             =   4365
      Width           =   1815
   End
   Begin VB.Label lblField 
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
      Index           =   54
      Left            =   9480
      TabIndex        =   47
      Top             =   3885
      Width           =   1815
   End
   Begin VB.Label lblField 
      Alignment       =   1  'Right Justify
      Caption         =   "Completion date:"
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
      Index           =   53
      Left            =   9480
      TabIndex        =   46
      Top             =   3405
      Width           =   1815
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
    fncPerson_First = 0
    fncPerson_Nickname
    fncPerson_Middle
    fncPerson_Last
    fncPerson_CellPhone
    fncPerson_Email
    fncPerson_DateOfBirth
    fncPerson_DateOfDeath

    'Person #2
    '(Same as above, but offset by PersonOffset)

    'Common
    fncUnused = 30
    fncIPTE
    fncMailingAddress_Street
    fncMailingAddress_City
    fncMailingAddress_State
    fncMailingAddress_ZipCode
    fncHomePhone
    fncWhichPhoneIsBest     'Uses three checkboxes,
    fncPLACEHOLDER1         '  hence the
    fncPLACEHOLDER2         '  placeholders
    fncNumApptSlots
    fncReminderCallAlways
    fncOldestYearFiled
    fncNewestYearFiled
    fncMailingListMode
    fncNotes1
    fncNotes2

    'TaxReturn #1
    fncTaxReturn_InboxType = 50
    fncTaxReturn_Status
    fncTaxReturn_AppointmentDate    'Not really a database field
    fncTaxReturn_CompletionDate
    fncTaxReturn_MinutesToComplete
    fncTaxReturn_FeeTotal
    fncTaxReturn_FeeOwed
    fncTaxReturn_ReleasedBeforePayment
    fncTaxReturn_ResultAGI
    fncTaxReturn_ResultFederal
    fncTaxReturn_ResultStatesCombined
    fncTaxReturn_StateList
    fncTaxReturn_EFiled
    fncTaxReturn_FiledExtension

    'TaxReturn #2, #3, #4, etc.
    '(Same as above, but offset by TaxReturnOffset)
End Enum

Public PhoneNumberCount As Long     'Used to determine how many checkboxes belong to the fncWhichPhoneIsBest group
Public PersonOffset As Long, TaxReturnOffset As Long
Public NumberOfTaxReturnsVisible As Long
Public TRColumnOffset As Long

Private mChooserConfig() As typeChooserConfig
Private ShowFormMode As enumShowFormMode
Private PreviouslyMarkedIncomplete As Boolean
Private This As DBModelClient
Private DataChanged As Boolean





'#################################################################################
'Load / Show / Save / Unload
'#################################################################################

'EHT=Standard
Private Sub Form_Load()
On Error GoTo ERR_HANDLER

If FormLoadedAlready Then Err.Raise 1, , "Attempted to load a form that had already been loaded."
FormLoadedAlready = True

'Since we cannot have a public constant, we'll simply initialize a public variable instead
PhoneNumberCount = 3
PersonOffset = 20
TaxReturnOffset = 20
NumberOfTaxReturnsVisible = 3
TRColumnOffset = 0

'Initialize the config data for the slider controls
Dim a&, b&, cci&
ReDim mChooserConfig(2 * NumberOfTaxReturnsVisible)
With mChooserConfig(0)
    Set .Container = chsField(fncMailingListMode)
    .ContainerIndex = .Container.Index
    ReDim .Selections(3)
    For b = 0 To 3
        Set .Selections(b) = lblChooserChoice(cci)
        cci = cci + 1
    Next b
End With
For a = 0 To NumberOfTaxReturnsVisible - 1
    With mChooserConfig(1 + (a * 2))
        Set .Container = chsField(fncTaxReturn_InboxType + (a * TaxReturnOffset))
        .ContainerIndex = .Container.Index
        ReDim .Selections(2)
        For b = 0 To 2
            Set .Selections(b) = lblChooserChoice(cci)
            cci = cci + 1
        Next b
    End With
    With mChooserConfig(2 + (a * 2))
        Set .Container = chsField(fncTaxReturn_Status + (a * TaxReturnOffset))
        .ContainerIndex = .Container.Index
        ReDim .Selections(3)
        For b = 0 To 3
            Set .Selections(b) = lblChooserChoice(cci)
            cci = cci + 1
        Next b
    End With
Next a

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Load", Err
End Sub

'EHT=None
Friend Function ChooserConfig(ByVal Field As enumClientTaxReturnFieldNum) As typeChooserConfig
Dim a&
For a = 0 To UBound(mChooserConfig)
    If mChooserConfig(a).ContainerIndex = Field Then
        ChooserConfig = mChooserConfig(a)
        Exit Function
    End If
Next a
Err.Raise 1, , "Cannot find ChooserConfig #" & Field
End Function

'EHT=Cleanup2
Function Form_Show(vShowFormMode As enumShowFormMode, vClient As DBModelClient, Optional vReadOnly As Boolean, Optional ByVal vOwnerForm As Form, Optional vNewClientInputString As String) As Boolean
On Error GoTo ERR_HANDLER: Dim INCLEANUP As Boolean, HASERROR As Boolean

'vShowFormMode              can be fPost, fEdit, or fNew
'vClient                    the DBModelClient to open
'vReadOnly                  if True, changes to vClient will not be allowed
'vOwnerForm                 only specify this if it is not frmMain or one of the tab 'forms'
'Return value               True if the DBModelClient was changed in any way or if new client; False if Cancel button was used to close the form

Dim a As Long, Y As Long

'Copy some parameters to global for later access
ShowFormMode = vShowFormMode
Set This = vClient

'Basic form initialize
Select Case ShowFormMode
Case fNew
    Me.Caption = "New Client"
    btnSave.Caption = "&Create"
    ResizeFormByInnerScaleDimensions Me, 625, 633
Case fPost
    Me.Caption = "Client #" & This.ID & " - Post"
    btnSave.Caption = "&Post"
    ResizeFormByInnerScaleDimensions Me, 857, 633
Case fEdit
    Me.Caption = "Client #" & This.ID & " - Edit"
    btnSave.Caption = "&Save"
    ResizeFormByInnerScaleDimensions Me, 1073, 633
    lblShiftTRView(1).Visible = True
End Select
btnSave.Enabled = Not vReadOnly

'Populate the form with real data
If Not This.PopulateToForm_Client(Me) Then
    'An error occured, and the user was already notified, so just quit
    HASERROR = True: GoTo CLEANUP
End If
For a = 0 To NumberOfTaxReturnsVisible - 1
    'The leftmost position will always be current year, but the rest will be shifted by TRColumnOffset
    If a = 0 Then
        Y = FileToOpen_Year - a
    Else
        Y = FileToOpen_Year - a - TRColumnOffset
    End If
    If Not This.AttachedDatabase Is Nothing Then
        If Not This.PopulateToForm_TaxReturn(Me, (a = 0), Y, a) Then
            'An error occured, and the user was already notified, so just quit
            HASERROR = True: GoTo CLEANUP
        End If
    End If
Next a
DataChanged = False

'No changes allowed on read-only database
btnSave.Enabled = frmMain.NEWDATABASE.IsWriteable

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

CLEANUP: INCLEANUP = True
    If HASERROR Then Unload Me

Exit Function
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "Form_Show", Err, INCLEANUP: HASERROR = True: Resume CLEANUP
End Function

'EHT=Standard
Private Sub btnSave_Click()
On Error GoTo ERR_HANDLER

If Not btnSave.Enabled Then Exit Sub

If This.PopulateFromForm(Me) Then
    If ShowFormMode <> fNew Then
        'We only set the changed flag if we're editing/posting a client that is actually in the database
        'New clients are not yet, so let the calling function handle this when/if it's added to the database
        frmMain.NEWDATABASE.SetDirty
    End If
    DataChanged = True
    Unload Me
End If

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

'EHT=ResumeNext
Private Sub Form_Resize()
On Error Resume Next

Dim sw!, sh!, tw!
sw = Me.ScaleWidth
sh = Me.ScaleHeight
lblChangeTabOrder.Move sw - lblChangeTabOrder.Width - 8, sh - lblChangeTabOrder.Height - 8
tw = btnSave.Width + btnCancel.Width + 23
btnSave.Left = (sw / 2) - (tw / 2)
btnCancel.Left = btnSave.Left + btnSave.Width + 23
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
        SetFocusWithoutErr btnSave
        btnSave_Click
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
'Field behavior
'#################################################################################

'##### Field labels #####

'EHT=ResumeNext
Private Sub lblField_Click(Index As Integer)
On Error Resume Next

Dim ctl As Control
Set ctl = txtField(Index)
If IsRealControl(ctl) Then
    SetFocusWithoutErr ctl
    SelectAll ctl
    Exit Sub
End If
Set ctl = chkField(Index)
If IsRealControl(ctl) Then
    SetFocusWithoutErr ctl
    Exit Sub
End If
Set ctl = chkgroupField(Index)
If IsRealControl(ctl) Then
    SetFocusWithoutErr ctl
    Exit Sub
End If
Set ctl = chsField(Index)
If IsRealControl(ctl) Then
    SetFocusWithoutErr ctl
    Exit Sub
End If
End Sub

'##### Text fields #####

'EHT=ResumeNext
Private Sub txtField_Change(Index As Integer)
On Error Resume Next

Select Case Index
Case fncPerson_Last, fncPerson_Last + PersonOffset
    Dim n$(1)
    n$(0) = LCase$(txtField(fncPerson_Last).Text)
    n$(1) = LCase$(txtField(fncPerson_Last + PersonOffset).Text)
    'If same last names, second one should be greyed out
    If Len(n$(1)) > 0 And (n$(1) = n$(0)) Then
        'Same = Grey
        txtField(fncPerson_Last + PersonOffset).ForeColor = &HC0C0C0
    Else
        'Different = Black
        txtField(fncPerson_Last + PersonOffset).ForeColor = vbWindowText
    End If

Case fncPerson_DateOfBirth, fncPerson_DateOfDeath, fncPerson_DateOfBirth + PersonOffset, fncPerson_DateOfDeath + PersonOffset
    UpdateDOBandDODtext

Case fncTaxReturn_ResultAGI, fncTaxReturn_ResultFederal, fncTaxReturn_ResultStatesCombined
    Dim v As Variant
    With txtField(Index)
        If ValidateTextbox(txtField(Index), v, False) Then
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
        Else
            'Error, so set it back to black
            .ForeColor = vbWindowText
        End If
    End With
End Select
End Sub
'EHT=ResumeNext
Private Sub txtField_GotFocus(Index As Integer)
On Error Resume Next

HilightControl Me, txtField(Index)

If ShowFormMode = fPost Then
    Select Case Index
    Case fncTaxReturn_FeeOwed
        If txtField(fncTaxReturn_FeeOwed).Text = "" Then
            Dim v As Long
            If ValidateChooser(ChooserConfig(fncTaxReturn_InboxType), v) Then
                If v = itDroppedOff Or v = itMailedIn Then
                    txtField(fncTaxReturn_FeeOwed).Text = txtField(fncTaxReturn_FeeTotal).Text
                End If
            End If
        End If
    Case fncTaxReturn_CompletionDate
        If txtField(fncTaxReturn_CompletionDate).Text = "" Then
            FieldToTextbox txtField(fncTaxReturn_CompletionDate), Date
        End If
    End Select
End If
End Sub
'EHT=ResumeNext
Private Sub txtField_LostFocus(Index As Integer)
On Error Resume Next

ClearControlHilight Me
ValidateTextbox txtField(Index), 0
End Sub

'##### Checkbox fields #####

'EHT=ResumeNext
Private Sub chkField_Click(Index As Integer)
On Error Resume Next

CheckboxClick chkField(Index)
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
ValidateCheckbox chkField(Index), False
End Sub

'##### Grouped checkbox fields #####

'EHT=ResumeNext
Private Sub chkgroupField_Click(Index As Integer)
On Error Resume Next

If (Index >= fncWhichPhoneIsBest) And (Index < (fncWhichPhoneIsBest + PhoneNumberCount)) Then
    'If user just checked a box in a gorup, uncheck the other ones
    If chkgroupField(Index).Value = vbChecked Then
        lblField(Index).Visible = True
        Dim a&
        For a = fncWhichPhoneIsBest To fncWhichPhoneIsBest + PhoneNumberCount - 1
            If a <> Index Then chkgroupField(a).Value = vbUnchecked
        Next a
    Else
        lblField(Index).Visible = False
    End If
End If
End Sub
'EHT=ResumeNext
Private Sub chkgroupField_GotFocus(Index As Integer)
On Error Resume Next

HilightControl Me, chkgroupField(Index)
End Sub
'EHT=ResumeNext
Private Sub chkgroupField_LostFocus(Index As Integer)
On Error Resume Next

ClearControlHilight Me
If (Index >= fncWhichPhoneIsBest) And (Index < (fncWhichPhoneIsBest + PhoneNumberCount)) Then
    ValidateCheckboxGroup chkgroupField, fncWhichPhoneIsBest, PhoneNumberCount, 0
End If
End Sub

'##### Chooser fields #####

'EHT=ResumeNext
Private Sub lblChooserChoice_Click(Index As Integer)
On Error Resume Next

Dim lbl As Label
Set lbl = lblChooserChoice(Index)
SetFocusWithoutErr lbl.Container: DoEvents
ChooserClick ChooserConfig(lbl.Container.Index), lbl
End Sub
'EHT=ResumeNext
Private Sub chsField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
Case vbKeyUp, vbKeyDown
    ChooserMove ChooserConfig(Index), (KeyCode = vbKeyDown)
End Select
End Sub
'EHT=ResumeNext
Private Sub chsField_GotFocus(Index As Integer)
On Error Resume Next

HilightControl Me, chsField(Index)
End Sub
'EHT=ResumeNext
Private Sub chsField_LostFocus(Index As Integer)
On Error Resume Next

ClearControlHilight Me
ValidateChooser ChooserConfig(Index), 0
End Sub

'##### Other controls that should be hilighted when focused #####

'EHT=ResumeNext
Private Sub btnSave_GotFocus()
On Error Resume Next

HilightControl Me, btnSave
End Sub
'EHT=ResumeNext
Private Sub btnSave_LostFocus()
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
Private Sub lblSwapPersons_Click(Index As Integer)
On Error GoTo ERR_HANDLER

SwitchTextboxValues fncPerson_First, fncPerson_First + PersonOffset
SwitchTextboxValues fncPerson_Nickname, fncPerson_Nickname + PersonOffset
SwitchTextboxValues fncPerson_Middle, fncPerson_Middle + PersonOffset
SwitchTextboxValues fncPerson_Last, fncPerson_Last + PersonOffset
SwitchTextboxValues fncPerson_Email, fncPerson_Email + PersonOffset
SwitchTextboxValues fncPerson_DateOfBirth, fncPerson_DateOfBirth + PersonOffset
SwitchTextboxValues fncPerson_DateOfDeath, fncPerson_DateOfDeath + PersonOffset
SwitchTextboxValues fncPerson_CellPhone, fncPerson_CellPhone + PersonOffset
SwitchCheckboxGroupValues fncWhichPhoneIsBest + 1, fncWhichPhoneIsBest + 2

SetFocusWithoutErr txtField(fncPerson_First)

Exit Sub
ERR_HANDLER: UNHANDLEDERROR MOD_NAME, "lblSwapPersons_Click", Err
End Sub

'EHT=None
Private Sub SwitchTextboxValues(txt1 As Integer, txt2 As Integer)
Dim v$
v$ = txtField(txt1).Text
txtField(txt1).Text = txtField(txt2).Text
txtField(txt2).Text = v$
End Sub

'EHT=None
Private Sub SwitchCheckboxGroupValues(chk1 As Integer, chk2 As Integer)
Dim v As Integer
v = chkgroupField(chk1).Value
chkgroupField(chk1).Value = chkgroupField(chk2).Value
chkgroupField(chk2).Value = v
End Sub





'#################################################################################
'Shifting of the tax return view to past years
'#################################################################################

'EHT=ResumeNext
Private Sub lblShiftTRView_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Dim a As Long

If Index = 0 Then
    TRColumnOffset = TRColumnOffset - 1
    If TRColumnOffset < 0 Then TRColumnOffset = 0
Else
    TRColumnOffset = TRColumnOffset + 1
End If
lblShiftTRView(0).Visible = (TRColumnOffset > 0)
imgTRDivider.Visible = (TRColumnOffset > 0)

For a = 1 To NumberOfTaxReturnsVisible - 1
    This.PopulateToForm_TaxReturn Me, False, FileToOpen_Year - a - TRColumnOffset, a
Next a
End Sub





'#################################################################################
'Calculations for DOB and DOD
'#################################################################################

'EHT=Standard
Sub UpdateDOBandDODtext()
On Error GoTo ERR_HANDLER

Dim a&, DOB&, DOD&
For a = 0 To 1
    lblDOBCalc(a).Visible = False
    lblDODCalc(a).Visible = False
    If ValidateTextbox(txtField(fncPerson_DateOfBirth + (a * PersonOffset)), DOB, False) And ValidateTextbox(txtField(fncPerson_DateOfDeath + (a * PersonOffset)), DOD, False) Then
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
            lblDODCalc(a).Visible = True
        ElseIf DOB <> NullLong Then
            If DOB <= Date Then
                lblDOBCalc(a).Caption = CalculateAge(DOB, Date) & "yr old today"
            Else
                lblDOBCalc(a).Caption = "ERR"
            End If
            lblDOBCalc(a).Visible = True
        End If
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
