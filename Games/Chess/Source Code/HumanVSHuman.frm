VERSION 5.00
Begin VB.Form frmChessMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   Caption         =   "Chess"
   ClientHeight    =   10740
   ClientLeft      =   1065
   ClientTop       =   1290
   ClientWidth     =   14565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "HumanVSHuman.frx":0000
   ScaleHeight     =   10740
   ScaleWidth      =   14565
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   13200
      TabIndex        =   116
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8160
      TabIndex        =   113
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   615
      Left            =   0
      TabIndex        =   112
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton cmdLoadGame 
      Caption         =   "Load Game"
      Height          =   615
      Left            =   5520
      TabIndex        =   111
      Top             =   0
      Width           =   2655
   End
   Begin VB.CommandButton cmdSaveGame 
      Caption         =   "Save Game"
      Height          =   615
      Left            =   2760
      TabIndex        =   110
      Top             =   0
      Width           =   2775
   End
   Begin VB.Timer BottomTimer 
      Interval        =   1000
      Left            =   12000
      Top             =   9120
   End
   Begin VB.Timer TopTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12000
      Top             =   720
   End
   Begin VB.TextBox BottomSec 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12360
      TabIndex        =   105
      Text            =   "00"
      Top             =   9600
      Width           =   615
   End
   Begin VB.TextBox BottomMin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11520
      TabIndex        =   104
      Text            =   "60"
      Top             =   9600
      Width           =   615
   End
   Begin VB.TextBox TOPsec 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12360
      MousePointer    =   1  'Arrow
      TabIndex        =   103
      Text            =   "00"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox TopMIN 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11400
      TabIndex        =   102
      Text            =   "60"
      Top             =   1080
      Width           =   615
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "HumanVSHuman.frx":212E1
      ScaleHeight     =   465
      ScaleWidth      =   11025
      TabIndex        =   101
      Top             =   9720
      Width           =   11055
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      Picture         =   "HumanVSHuman.frx":D0403
      ScaleHeight     =   465
      ScaleWidth      =   11025
      TabIndex        =   98
      Top             =   600
      Width           =   11055
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   10320
      Picture         =   "HumanVSHuman.frx":17F525
      ScaleHeight     =   8985
      ScaleWidth      =   705
      TabIndex        =   99
      Top             =   960
      Width           =   735
      Begin VB.PictureBox Picture5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   100
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   0
      Picture         =   "HumanVSHuman.frx":1E2663
      ScaleHeight     =   8745
      ScaleWidth      =   705
      TabIndex        =   96
      Top             =   960
      Width           =   735
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   735
         TabIndex        =   97
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":2457A1
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   95
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":245919
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   94
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":245C80
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   93
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":245FCE
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   92
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":2465E4
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   91
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":246B46
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   90
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":246E94
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   89
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   32
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":2471FB
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   88
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":247373
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   87
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":247575
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   86
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":247777
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   85
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":247979
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   84
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":247B7B
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   83
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":247D7D
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   82
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":247F7F
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   81
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":248181
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   80
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":248383
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":248423
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   14
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":2484C3
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   13
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":248563
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   12
      Top             =   6480
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":248603
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":2486A3
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":248743
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":2487E3
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   5640
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":248883
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   7320
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":248973
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   7320
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   12000
      Picture         =   "HumanVSHuman.frx":248C3A
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   8160
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":248D38
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   8160
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":249295
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   8160
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   12720
      Picture         =   "HumanVSHuman.frx":2495A0
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   8160
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   13440
      Picture         =   "HumanVSHuman.frx":24969E
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   7320
      Width           =   735
   End
   Begin VB.PictureBox BlackxPieces 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   11280
      Picture         =   "HumanVSHuman.frx":249965
      ScaleHeight     =   855
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   7320
      Width           =   735
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   16
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   17
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   18
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   19
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   20
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   21
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   22
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   23
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   24
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   10
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   25
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   26
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   12
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   27
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   13
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   15
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   30
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   16
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   31
      Top             =   7560
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   17
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   32
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   18
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   33
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   19
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   34
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   20
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   35
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   21
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   36
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   22
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   37
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   23
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   38
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   24
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   39
      Top             =   6480
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   25
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   40
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   26
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   41
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   27
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   42
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   28
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   43
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   29
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   44
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   30
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   45
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   31
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   46
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   32
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   47
      Top             =   5400
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   33
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   48
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   34
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   49
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   35
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   50
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   36
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   51
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   37
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   52
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   38
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   53
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   39
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   54
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   40
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   55
      Top             =   4320
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   41
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   56
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   42
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   57
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   43
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   58
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   44
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   59
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   45
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   60
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   46
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   61
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   47
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   62
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   48
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   63
      Top             =   3240
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   49
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   64
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   50
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   65
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   51
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   66
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   52
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   67
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   53
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   68
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   54
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   69
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   55
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   70
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   56
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   71
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   57
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   72
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   58
      Left            =   1920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   73
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   59
      Left            =   3120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   74
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   60
      Left            =   4320
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   75
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   61
      Left            =   5520
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   76
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   62
      Left            =   6720
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   77
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   63
      Left            =   7920
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   78
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Grid 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   64
      Left            =   9120
      ScaleHeight     =   1035
      ScaleWidth      =   1155
      TabIndex        =   79
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox picWhtPrison 
      BackColor       =   &H00E0E0E0&
      Height          =   3615
      Left            =   11160
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   114
      Top             =   1920
      Width           =   3135
   End
   Begin VB.PictureBox picBlckPrison 
      BackColor       =   &H00404040&
      Height          =   3615
      Left            =   11160
      ScaleHeight     =   3555
      ScaleWidth      =   3075
      TabIndex        =   115
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   109
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   108
      Top             =   9240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   107
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   106
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmChessMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GameOver(WinningPlayer As Integer, WinType As Integer)

If WinType = 1 Then ' Checkmate
 MsgBox "Winner is:", vbOKOnly, "Checkmate"
End If

If WinType = 2 Then ' Time Over
MsgBox "Player", vbOKOnly, "Time Over"
End If

End Function

Private Sub cmdDelLoad_Click()

Dim Stro As String

    Kill App.Path + "\SaveGames\" + LoadList.FileName

LoadList.Refresh




End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLoad02_Click()
    Call LoadList_DblClick
End Sub

Private Sub cmdNewGame_Click()
    
   Dim StoreTurn As Boolean
    
        
   Dim Counter As Integer
   
   'LoadList.Visible = True
   
   
   Let Counter = 0
    

   
   Call LoadGame2("Default.chsX")
   

    
End Sub

Private Sub cmdLoadGame_Click()
    
    

frmFileMenu.Show
frmFileMenu.Refresh
frmFileMenu.flstFileList.Refresh
frmFileMenu.cmdSaveLoad02.Caption = "Load"

    
    
    

    
End Sub


Private Sub cmdSaveGame_Click()
   
'frmFileMenu.
frmFileMenu.Show
frmFileMenu.cmdSaveLoad02.Caption = "Save"
   
   'LoadList.Visible = False
   'Call SaveGame(TopMIN.Text, TOPsec.Text, BottomMin.Text, BottomSec.Text)


End Sub


Private Sub LoadList_DblClick()
    
       
Dim StoreTurn As Boolean

       
   LoadList.Visible = False
  

   Dim Counter As Integer
   
   Let Counter = 0
    
    Erase Board
    Erase UnitLocations
   
   
   
   Call LoadGame(LoadList.FileName)
   
 If Turn = True Then
        Let StoreTurn = True
 Else
        Let StoreTurn = False
 End If
 
   
   Do
   
   Counter = Counter + 1
   If Counter = 33 Then Exit Do
         
   If MyColumn(Counter) = 0 Or MyRow(Counter) = 0 Then
   
        BlackxPieces(Counter).Enabled = False
        
        
        BlackxPieces(Counter).Left = PrisonerLocations(1, Counter)
        BlackxPieces(Counter).Top = PrisonerLocations(2, Counter)
        BlackxPieces(Counter).BackColor = frmChessMain.BackColor
         
   
   Else
   
    Call XMoveUnit(Counter, MyRow(Counter), MyColumn(Counter), True)
   
   End If
   
   TopMIN.Text = Timings(1, 1)
   TOPsec.Text = Timings(2, 1)
   
   BottomMin.Text = Timings(1, 2)
   BottomSec.Text = Timings(2, 2)
   
   
   
   Loop
   
   
   
 If StoreTurn = False Then
            
            Let Turn = False
            TopTimer.Enabled = True
            BottomTimer.Enabled = False
        
   ElseIf StoreTurn = True Then
         
            Let Turn = True
            TopTimer.Enabled = False
            BottomTimer.Enabled = True
End If
  
   
   

End Sub

Private Sub Command1_Click()
Call MateCheck

End Sub

Private Sub Form_Load()
Debugger.Show

Dim Counter As Byte

Counter = 0

Do

Counter = Counter + 1
If Counter = 33 Then Exit Do

PrisonerLocations(1, Counter) = BlackxPieces(Counter).Left
PrisonerLocations(2, Counter) = BlackxPieces(Counter).Top



Loop


'LoadList.Visible = False

cmdNewGame_Click



End Sub


Private Sub BlackxPieces_Click(Index As Integer)


If Not PieceIndex = 0 Then
    If ScoutPlace(PieceIndex, UnitLocations(1, Index), UnitLocations(2, Index)) = "Enemy Occupied" Then
        Call Grid_MouseUp(gridTolinear(UnitLocations(1, Index), UnitLocations(2, Index)), 2, 0, 0, 0)
        GoTo Skip
    End If
End If

        
   
   



If Turn = True Then
If Index >= 17 Then
MsgBox "This is Player 1's turn"
GoTo Skip
End If
End If

If Turn = False Then
If Index < 17 Then
MsgBox "This is Player 2's turn"
GoTo Skip
End If
End If





If Not PieceIndex = 0 Then
End If


  
Dim CounterA9 As Integer
CounterA9 = 1

Do Until CounterA9 = 33
    
    BlackxPieces(CounterA9).BorderStyle = 0
       
    CounterA9 = CounterA9 + 1

Loop



  
PieceIndex = Index  ' Assign the Clicked Chess Piece as the current Piece to move.

BlackxPieces(Index).BorderStyle = 1     ' Higlight the currently Selected Piece



  
'--------------------------------------------------------------------
' Compute Possible Moves
'--------------------------------------------------------------------


Call ClearHilighting




Call HilightMoves(PieceIndex, UnitLocations(1, PieceIndex), UnitLocations(2, PieceIndex))





Skip:


End Sub



Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Grid_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)



Dim SID As Integer
'SID = TextBox.Text



If (PieceIndex = 0) And (Button = 1) And (Shift = 1) Then ' Left Mouse Button

GridIndex = Index
'Call ClearHilighting
'Call AssignMoves(SID, GetRow(GridIndex), GetColumn(GridIndex))
'Call HilightMoves(SID, GetRow(GridIndex), GetColumn(GridIndex))
GoTo Skip

End If


If Button = 2 Then ' Right Click

    If PieceIndex = 0 Then
        MsgBox "Please Select a Unit to Move.", vbOKOnly, "No Unit Selected"
        GoTo Skip
        
    Else
        GridIndex = Index
        Call XMoveUnit(PieceIndex, GetRow(GridIndex), GetColumn(GridIndex), False)
        
    End If




End If


Skip:
    
  

End Sub







Private Sub TopTimer_Timer()

If TOPsec.Text = 0 And TopMIN.Text = 60 Then
    Let TopMIN.Text = 59
    Let TOPsec.Text = 59


ElseIf TopMIN.Text = 0 And TOPsec.Text = 1 Then
    Let TOPsec.Text = TOPsec.Text - 1
    Let TopMIN.Text = 0
    TopTimer.Enabled = False
    
ElseIf TopMIN.Text = 0 And TOPsec.Text - 1 > 0 Then
    Let TOPsec.Text = TOPsec.Text - 1

ElseIf TOPsec.Text = 0 Then
    Let TOPsec.Text = 59
    Let TopMIN.Text = TopMIN.Text - 1
    
ElseIf TOPsec.Text = 0 And TopMIN.Text = 0 Then
    TopTimer.Enabled = False
    
ElseIf TOPsec.Text > 0 And TopMIN.Text > 0 Then
    Let TOPsec.Text = TOPsec.Text - 1
    

End If


End Sub

Private Sub BottomTimer_Timer()



If BottomSec.Text = 0 And BottomMin.Text = 60 Then
    Let BottomMin.Text = 59
    Let BottomSec.Text = 59


ElseIf BottomMin.Text = 0 And BottomSec.Text = 1 Then
    Let BottomSec.Text = BottomSec.Text - 1
    Let BottomMin.Text = 0
    BottomTimer.Enabled = False
    
ElseIf BottomMin.Text = 0 And BottomSec.Text - 1 > 0 Then
    Let BottomSec.Text = BottomSec.Text - 1

ElseIf BottomSec.Text = 0 Then
    Let BottomSec.Text = 59
    Let BottomMin.Text = BottomMin.Text - 1
    
ElseIf BottomSec.Text = 0 And BottomMin.Text = 0 Then
    BottomTimer.Enabled = False
    
ElseIf BottomSec.Text > 0 And BottomMin.Text > 0 Then
    Let BottomSec.Text = BottomSec.Text - 1
    

End If


End Sub

Private Sub Write_Click()

Call AIMoveFunction

CounterX1 = 0

Console.Print "LBOund"; LBound(AIMoves, 2)
Console.Print "UBOund"; UBound(AIMoves, 2)

Do Until CounterX1 = UBound(AIMoves, 2)

CounterX1 = CounterX1 + 1

Console.Print "Unit"; AIMoves(1, CounterX1)
Console.Print "Row"; AIMoves(2, CounterX1)
Console.Print "Column"; AIMoves(3, CounterX1)

Loop
'XMoveUNit(unitID,Row,COlumn)
'Number = Int(Rnd * (6 - 1 + 1)) + 1
Val1 = (Int((UBound(AIMoves, 2) - LBound(AIMoves, 2) + 1) * Rnd() + LBound(AIMoves, 2)))

Call XMoveUnit(AIMoves(1, Val1), AIMoves(2, Val1), AIMoves(3, Val1), True)



End Sub
