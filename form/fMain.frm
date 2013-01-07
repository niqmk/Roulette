VERSION 5.00
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "ROULETTE SYSTEM"
   ClientHeight    =   8775
   ClientLeft      =   105
   ClientTop       =   -285
   ClientWidth     =   6750
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   6750
   Begin VB.CommandButton cdEngage 
      Caption         =   "Engage"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   167
      TabStop         =   0   'False
      Top             =   5040
      Width           =   735
   End
   Begin VB.Timer trDelay 
      Left            =   3600
      Top             =   960
   End
   Begin VB.CommandButton cdLow 
      Caption         =   "Low"
      Height          =   375
      Left            =   5880
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton cdOdd 
      Caption         =   "Odd"
      Height          =   375
      Left            =   5880
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cdRed 
      Caption         =   "Red"
      Height          =   375
      Left            =   5880
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cdHigh 
      Caption         =   "High"
      Height          =   375
      Left            =   5880
      TabIndex        =   163
      TabStop         =   0   'False
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cdEven 
      Caption         =   "Even"
      Height          =   375
      Left            =   5880
      TabIndex        =   161
      TabStop         =   0   'False
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cdBlack 
      Caption         =   "Black"
      Height          =   375
      Left            =   5880
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txPassword 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      ScrollBars      =   1  'Horizontal
      TabIndex        =   150
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cdClear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   149
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cdHistory 
      Caption         =   "History"
      Height          =   375
      Left            =   5880
      TabIndex        =   148
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cdTable2 
      Caption         =   "Table 2"
      Height          =   375
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cdTable1 
      Caption         =   "Table 1"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   146
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame frBet 
      BackColor       =   &H00F0F0F0&
      Caption         =   "BET"
      Height          =   1455
      Left            =   3600
      TabIndex        =   142
      Top             =   7200
      Width           =   3015
      Begin VB.Timer trBlink 
         Index           =   6
         Left            =   120
         Top             =   600
      End
      Begin VB.Timer trBlink 
         Index           =   5
         Left            =   2520
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   4
         Left            =   2040
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   3
         Left            =   1560
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   2
         Left            =   1080
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   1
         Left            =   600
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   0
         Left            =   120
         Top             =   120
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   157
         Top             =   960
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   156
         Top             =   600
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   155
         Top             =   240
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2520
         TabIndex        =   154
         Top             =   600
         Width           =   375
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   145
         Top             =   960
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   144
         Top             =   600
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbBet 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   143
         Top             =   240
         Width           =   1815
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox pcGreen 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5280
      Picture         =   "fMain.frx":08CA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   141
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pcRed 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5280
      Picture         =   "fMain.frx":0EEB
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   140
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "36"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   36
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   139
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   35
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   138
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "34"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   32
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   31
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   30
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   133
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   29
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   28
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   27
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   26
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   129
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   25
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   128
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   127
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   126
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   125
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   20
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H008080FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cdRoll 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox pcNoneFocus 
      Height          =   375
      Left            =   4800
      Picture         =   "fMain.frx":15E8
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcCrossFocus 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   4200
      Picture         =   "fMain.frx":179C
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcTickFocus 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "fMain.frx":1A5E
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5880
      TabIndex        =   70
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cdBox 
      Caption         =   "Score Board"
      Height          =   375
      Left            =   3600
      TabIndex        =   68
      Top             =   6120
      Width           =   2175
   End
   Begin VB.PictureBox pcMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   5760
      Picture         =   "fMain.frx":1CA4
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   66
      Top             =   120
      Width           =   420
   End
   Begin VB.PictureBox pcExit 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   6240
      Picture         =   "fMain.frx":2616
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   65
      Top             =   120
      Width           =   420
   End
   Begin VB.PictureBox pcNone 
      Height          =   375
      Left            =   4800
      Picture         =   "fMain.frx":2F88
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcCross 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   4200
      Picture         =   "fMain.frx":3129
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcTick 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "fMain.frx":341C
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   8280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcTickLast 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   3600
      Picture         =   "fMain.frx":3653
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcCrossLast 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   4200
      Picture         =   "fMain.frx":3877
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   152
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcNoneLast 
      Height          =   375
      Left            =   4800
      Picture         =   "fMain.frx":3AF2
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame frMain 
      BackColor       =   &H00F0F0F0&
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   102
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   101
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   100
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   99
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   98
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   97
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   96
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   95
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   94
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   93
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   92
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   91
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   90
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   89
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   88
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   87
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   86
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   85
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   58
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   57
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   56
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   55
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   54
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   53
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   52
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   51
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   50
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   49
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   48
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   47
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   42
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   40
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   39
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   37
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   36
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   35
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   34
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   33
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   30
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   29
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   28
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   27
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   26
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   25
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   24
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   23
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   22
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   21
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   32
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcHigh 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox pcEven 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1920
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   31
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1080
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   19
         Left            =   120
         TabIndex        =   79
         Top             =   7440
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   18
         Left            =   120
         TabIndex        =   78
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   17
         Left            =   120
         TabIndex        =   77
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   16
         Left            =   120
         TabIndex        =   76
         Top             =   6360
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   15
         Left            =   120
         TabIndex        =   75
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   14
         Left            =   120
         TabIndex        =   74
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   13
         Left            =   120
         TabIndex        =   16
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   12
         Left            =   120
         TabIndex        =   64
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   11
         Left            =   120
         TabIndex        =   15
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   10
         Left            =   120
         TabIndex        =   14
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   9
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   8
         Left            =   120
         TabIndex        =   63
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbRoll 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0F0&
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
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HIGH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   4
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EVEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   510
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BLACK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   660
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ROLL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   525
      End
      Begin VB.Line lnBorder 
         Index           =   4
         X1              =   0
         X2              =   3360
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Label lbDragDrop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   6480
      TabIndex        =   166
      ToolTipText     =   "No Delay"
      Top             =   720
      Width           =   135
   End
   Begin VB.Label lbDragDrop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   3600
      TabIndex        =   165
      ToolTipText     =   "Delay"
      Top             =   720
      Width           =   135
   End
   Begin VB.Label lbCount 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   158
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   84
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   15
      Index           =   4
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   8895
      Index           =   3
      Left            =   6720
      TabIndex        =   82
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   15
      Index           =   2
      Left            =   0
      TabIndex        =   81
      Top             =   8760
      Width           =   6735
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   8895
      Index           =   1
      Left            =   0
      TabIndex        =   80
      Top             =   -120
      Width           =   15
   End
   Begin VB.Label lbCount 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   69
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00FF8080&
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "RS Quicky"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   62
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BBox As Boolean
Public BTable1 As Boolean
Public BTable2 As Boolean
Public BHistory As Boolean
Public BBlack As Boolean
Public BRed As Boolean
Public BEven As Boolean
Public BOdd As Boolean
Public BHigh As Boolean
Public BLow As Boolean

Private Const LStartButton As Long = &HFF00&
Private Const LStopButton As Long = &H8080FF
Private Const SStartButton As String = "Start"
Private Const SStopButton As String = "Stop"

Private IBlackTimer As Integer
Private IEvenTimer As Integer
Private IHighTimer As Integer
Private IMiddleTimer As Integer
Private IBLeftTimer As Integer
Private IELeftTimer As Integer
Private IHLeftTimer As Integer
Private SnX As Single
Private SnY As Single
Private BMove As Boolean
Private BCountDownFlag As Boolean
Private BWinFlag As Boolean
Private BLossFlag As Boolean
Private BEngage As Boolean

Public Sub CloseForm()
    Unload Me
End Sub

Public Sub HideAllBox()
    If BBox Then fBox.Hide
    If BTable1 Then fRoll1.Hide
    If BTable2 Then fRoll2.Hide
    If BHistory Then fHistory.Hide
    If BBlack Then fBlack.Hide
    If BRed Then fRed.Hide
    If BEven Then fEven.Hide
    If BOdd Then fOdd.Hide
    If BHigh Then fHigh.Hide
    If BLow Then fLow.Hide
End Sub

Public Sub EDAllBOx(Optional ByVal BEnable As Boolean = True)
    If BBox Then fBox.Enabled = BEnable
    If BTable1 Then fRoll1.Enabled = BEnable
    If BTable2 Then fRoll2.Enabled = BEnable
    If BHistory Then fHistory.Enabled = BEnable
    If BBlack Then fBlack.Enabled = BEnable
    If BRed Then fRed.Enabled = BEnable
    If BEven Then fEven.Enabled = BEnable
    If BOdd Then fOdd.Enabled = BEnable
    If BHigh Then fHigh.Enabled = BEnable
    If BLow Then fLow.Enabled = BEnable
End Sub

Public Sub ClearBox()
    mdApp.DeleteBoxHistory
    
    If BBlack Then fBlack.SetText
    If BRed Then fRed.SetText
    If BBlack Then fEven.SetText
    If BOdd Then fOdd.SetText
    If BHigh Then fHigh.SetText
    If BLow Then fLow.SetText
End Sub

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyDelete Then
        SetDelete
    ElseIf KeyCode = vbKeyF3 Then
        ChangePassword
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BMove = True
    
    SnX = X
    SnY = Y
    
    Me.MousePointer = MousePointerConstants.vbSizeAll
End Sub

Private Sub lbDragDrop_DblClick(Index As Integer)
    If (UBound(mdApp.INumberRollBoard) = 0) And (mdApp.INumberRollBoard(0) = mdApp.IBlank) Then Exit Sub
    
    Me.MousePointer = MousePointerConstants.vbHourglass
    
    Dim ICounter As Integer
    
    Dim SValue As String
    
    SValue = ""
    
    For ICounter = LBound(mdApp.INumberRollBoard) To UBound(mdApp.INumberRollBoard)
        If Not (Trim(SValue) = "") Then SValue = SValue & ","
        
        SValue = SValue & CStr(mdApp.INumberRollBoard(ICounter))
    Next ICounter
    
    Open mdGeneral.SPath & mdApp.STitle & ".txt" For Output As #1
        Print #1, SValue
    Close #1
    
    Me.MousePointer = MousePointerConstants.vbIconPointer
End Sub

Private Sub lbDragDrop_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.GetFormat(vbCFText) Then
        Dim SValue As String
        
        SValue = Data.GetData(vbCFText)
        
        Dim ICounter As Integer
        
        Dim SData() As String
        
        SData = Split(SValue, ",")
        
        If (UBound(SData) + UBound(mdApp.INumberRollBoard) + 2) > mdApp.IRollBaseDec Then
            MsgBox "Too Many Imported Data", vbCritical, mdApp.STitle
            
            Exit Sub
        End If
        
        Me.MousePointer = MousePointerConstants.vbHourglass
        Me.Enabled = False
        EDAllBOx False
        
        If Index = 0 Then
            For ICounter = LBound(SData) To UBound(SData)
                If IsNumeric(Trim(SData(ICounter))) And (Not Trim(SData(ICounter)) = "00") Then
                    SetRoll CInt(SData(ICounter))
                    
                    mdGeneral.Sleep 1000000
                End If
            Next ICounter
        ElseIf Index = 1 Then
            For ICounter = LBound(SData) To UBound(SData)
                If IsNumeric(Trim(SData(ICounter))) And (Not Trim(SData(ICounter)) = "00") Then
                    SetRoll CInt(SData(ICounter)), False
                    
                    DoEvents
                End If
            Next ICounter
            
            ShowAfterNoAlert
        End If
        
        Me.MousePointer = MousePointerConstants.vbIconPointer
        Me.Enabled = True
        EDAllBOx
    End If
End Sub

Private Sub lbTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SnXDrag, SnYDrag As Single

    If BMove Then
        SnXDrag = Me.Left + X - SnX
        SnYDrag = Me.Top + Y - SnY
        
        Me.Left = SnXDrag
        Me.Top = SnYDrag
    End If
End Sub

Private Sub lbTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SnXDrag, SnYDrag As Single

    If BMove Then
        SnXDrag = Me.Left + X - SnX
        SnYDrag = Me.Top + Y - SnY
        
        Me.Left = SnXDrag
        Me.Top = SnYDrag
        
        BMove = False
    End If
    
    Me.MousePointer = MousePointerConstants.vbDefault
End Sub

Private Sub lbTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fMain = Nothing
End Sub

Private Sub txPassword_LostFocus()
    Me.txPassword.Visible = False
End Sub

Private Sub txPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SetClear
    End If
End Sub

Private Sub cdClear_Click()
    Me.txPassword.Text = ""
    Me.txPassword.Visible = True
    Me.txPassword.SetFocus
End Sub

Private Sub cdRoll_Click(Index As Integer)
    SetRoll Index
End Sub

Private Sub cdBox_Click()
    If BBox Then
        fBox.CloseForm
        
        BBox = False
    Else
        BCountDownFlag = False
    
        fBox.Show vbModeless, Me
        
        mdApp.SetWin
        mdApp.SetLoss
        
        fBox.lbWin.Caption = CStr(mdApp.CheckWin)
        fBox.lbLoss.Caption = CStr(mdApp.CheckLoss)
        
        Me.SetFocus
        
        BBox = True
    End If
End Sub

Private Sub cdDelete_Click()
    SetDelete
End Sub

Private Sub cdTable1_Click()
    If BTable1 Then
        fRoll1.CloseForm
        
        BTable1 = False
    Else
        fRoll1.Show vbModeless, Me
        
        BTable1 = True
    End If
End Sub

Private Sub cdTable2_Click()
    If BTable2 Then
        fRoll2.CloseForm
        
        BTable2 = False
    Else
        fRoll2.Show vbModeless, Me
        
        BTable2 = True
    End If
End Sub

Private Sub cdHistory_Click()
    If BHistory Then
        fHistory.CloseForm
        
        BHistory = False
    Else
        fHistory.Show vbModeless, Me
        
        BHistory = True
    End If
End Sub

Private Sub cdBlack_Click()
    If BBlack Then
        fBlack.CloseForm
        
        BBlack = False
    Else
        fBlack.Show vbModeless, Me
        
        BBlack = True
    End If
End Sub

Private Sub cdRed_Click()
    If BRed Then
        fRed.CloseForm
        
        BRed = False
    Else
        fRed.Show vbModeless, Me
        
        BRed = True
    End If
End Sub

Private Sub cdEven_Click()
    If BEven Then
        fEven.CloseForm
        
        BEven = False
    Else
        fEven.Show vbModeless, Me
        
        BEven = True
    End If
End Sub

Private Sub cdOdd_Click()
    If BOdd Then
        fOdd.CloseForm
        
        BOdd = False
    Else
        fOdd.Show vbModeless, Me
        
        BOdd = True
    End If
End Sub

Private Sub cdHigh_Click()
    If BHigh Then
        fHigh.CloseForm
        
        BHigh = False
    Else
        fHigh.Show vbModeless, Me
        
        BHigh = True
    End If
End Sub

Private Sub cdLow_Click()
    If BLow Then
        fLow.CloseForm
        
        BLow = False
    Else
        fLow.Show vbModeless, Me
        
        BLow = True
    End If
End Sub

Private Sub cdEngage_Click()
    If BEngage Then
        Me.cdEngage.Caption = SStopButton
        Me.cdEngage.BackColor = LStopButton
        
        BEngage = False
    Else
        Me.cdEngage.Caption = SStartButton
        Me.cdEngage.BackColor = LStartButton
        
        BEngage = True
    End If
End Sub

Private Sub pcExit_Click()
    Unload Me
End Sub

Private Sub pcMin_Click()
    Me.WindowState = FormWindowStateConstants.vbMinimized
End Sub

Private Sub lbBet_Change(Index As Integer)
    If Not (Trim(Me.lbBet(Index).Caption) = "") Then
        If Trim(Me.lbBet(Index).Caption) = "0" Then
            Me.lbBet(Index).ForeColor = LGreen
        Else
            Me.lbBet(Index).ForeColor = LBlack
        End If
        
        If (Index = 4) Or (Index = 5) Or (Index = 6) Then
            Me.lbBet(Index).ForeColor = mdApp.LBlue
        End If
        
        Me.trBlink(Index).Enabled = True
    End If
End Sub

Private Sub trBlink_Timer(Index As Integer)
    Me.lbBet(Index).Visible = Not Me.lbBet(Index).Visible
    
    Select Case Index
        Case 0:
            If IBlackTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IBlackTimer = 250
            Else
                IBlackTimer = IBlackTimer + 250
            End If
        Case 1:
            If IEvenTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IEvenTimer = 250
            Else
                IEvenTimer = IEvenTimer + 250
            End If
        Case 2:
            If IHighTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IHighTimer = 250
            Else
                IHighTimer = IHighTimer + 250
            End If
        Case 3:
            If IMiddleTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IMiddleTimer = 250
            Else
                IMiddleTimer = IMiddleTimer + 250
            End If
        Case 4:
            If IBLeftTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IBLeftTimer = 350
            Else
                IBLeftTimer = IBLeftTimer + 250
            End If
        Case 5:
            If IELeftTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IELeftTimer = 350
            Else
                IELeftTimer = IELeftTimer + 250
            End If
        Case 6:
            If IHLeftTimer >= 2000 Then
                Me.trBlink(Index).Enabled = False
                Me.lbBet(Index).Visible = True
                
                IHLeftTimer = 350
            Else
                IHLeftTimer = IHLeftTimer + 250
            End If
    End Select
End Sub

Private Sub SetInitial()
    Dim ICounter As Integer
    
    mdGeneral.CenterWindows Me, False

    mdApp.Init
    
    Me.txPassword.Visible = False
    Me.lbDragDrop(0).OLEDropMode = 1
    Me.lbDragDrop(1).OLEDropMode = 1
    
    For ICounter = 0 To 3
        Me.trBlink(ICounter).Interval = 250
        Me.trBlink(ICounter).Enabled = False
    Next ICounter
    
    For ICounter = 4 To Me.trBlink.Count - 1
        Me.trBlink(ICounter).Interval = 300
        Me.trBlink(ICounter).Enabled = False
    Next ICounter
    
    Me.trDelay.Interval = 3000
    Me.trDelay.Enabled = False
    
    IBlackTimer = 250
    IEvenTimer = 250
    IHighTimer = 250
    IMiddleTimer = 250
    IBLeftTimer = 350
    IELeftTimer = 350
    IHLeftTimer = 350
    IHLeftTimer = 350
    
    BBox = False
    BTable1 = False
    BTable2 = False
    BHistory = False
    BBlack = False
    BRed = False
    BEven = False
    BOdd = False
    BHigh = False
    BLow = False
    BCountDownFlag = False
    BEngage = True
    
    Me.cdEngage.Caption = SStartButton
    Me.cdEngage.BackColor = LStartButton
    
    If mdApp.CheckTableHistory = 1 Then
        Me.cdTable1.BackColor = mdApp.LLGreen
        Me.cdTable2.BackColor = mdApp.LLRed
    Else
        Me.cdTable1.BackColor = mdApp.LLRed
        Me.cdTable2.BackColor = mdApp.LLGreen
    End If
    
    SetRollBoard
    
    If UBound(mdApp.INumberRollBoard) = 0 Then
        If mdApp.INumberRollBoard(0) = IBlank Then
            Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard), "#,##0")
        Else
            Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard) + 1, "#,##0")
        End If
    Else
        Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard) + 1, "#,##0")
    End If
    Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
    
    If mdSecurity.GetGuest Then fTimer.Show , Me
End Sub

Private Sub SetRoll(ByVal IRoll As Integer, Optional ByVal BAlert As Boolean = True)
    If mdApp.ICounterRollDec = 0 Then Exit Sub
    
    mdApp.ICounterRollDec = mdApp.ICounterRollDec - 1
    
    If IRoll > mdApp.INumberMax Or IRoll < mdApp.INone Then Exit Sub
    
    If IRoll = mdApp.INone Then
        If mdApp.CheckZero > 1 Then Exit Sub
    End If
    
    SetBoardPattern IRoll, BAlert
    
    If BHistory Then fHistory.AddRoll IRoll
    
    SetBetText False, BAlert
    
    If BAlert Then
        SetRollBoard
        SetRollFocus
        SetRollHistory
    End If
    
    mdApp.SetHistory IRoll
    mdApp.SaveBoxHistory
    
    If BAlert Then
        Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard) + 1, "#,##0")
        Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
    End If
    
    If BBox Then
        If BCountDownFlag Then
            If BWinFlag Then
                If mdApp.CheckWin <= 0 Then
                    mdApp.ICounterRollDec = mdApp.IRollBaseDec
                    mdAPI.WriteValueRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROLL DEC", CStr(ICounterRollDec)
                    Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
                
                    BCountDownFlag = False
                    BWinFlag = False
                End If
            End If
            
            If BLossFlag Then
                If mdApp.CheckLoss <= 0 Then
                    mdApp.ICounterRollDec = mdApp.IRollBaseDec
                    mdAPI.WriteValueRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROLL DEC", CStr(ICounterRollDec)
                    Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
                
                    BCountDownFlag = False
                    BLossFlag = False
                End If
            End If
        Else
            If (mdApp.CheckWin <> 0) Or (mdApp.CheckLoss <> 0) Then
                If mdApp.CheckWin > 0 Then BWinFlag = True
                If mdApp.CheckLoss > 0 Then BLossFlag = True
                
                BCountDownFlag = True
            End If
        End If
        
        fBox.lbWin.Caption = CStr(mdApp.CheckWin)
        fBox.lbLoss.Caption = CStr(mdApp.CheckLoss)
    End If
    
    If mdApp.ICounterRollDec = 0 Then Me.CloseForm
End Sub

Private Sub SetBoardPattern(ByVal IRoll As Integer, Optional ByVal BAlert As Boolean = True)
    mdApp.SetNumberRollBoard IRoll
    
    ReDim Preserve mdApp.IRollBPatternHTS(UBound(mdApp.INumberRollBoard)) As Integer
    ReDim Preserve mdApp.IRollEPatternHTS(UBound(mdApp.INumberRollBoard)) As Integer
    ReDim Preserve mdApp.IRollHPatternHTS(UBound(mdApp.INumberRollBoard)) As Integer
    
    mdApp.IRollBPatternHTS(UBound(mdApp.IRollBPatternHTS)) = IBlank
    mdApp.IRollEPatternHTS(UBound(mdApp.IRollEPatternHTS)) = IBlank
    mdApp.IRollHPatternHTS(UBound(mdApp.IRollHPatternHTS)) = IBlank
    
    If mdApp.LColorPattern(IRoll) = mdApp.LRed Then
        mdApp.SetBlackRollBoard mdApp.ICross
        mdApp.SetBlackPattern mdApp.ICross
        
        mdApp.SetBlackBox LOSS
        mdApp.SetRedBox WIN
    ElseIf LColorPattern(IRoll) = mdApp.LBlack Then
        mdApp.SetBlackRollBoard mdApp.ITick
        mdApp.SetBlackPattern mdApp.ITick
        
        mdApp.SetBlackBox WIN
        mdApp.SetRedBox LOSS
    ElseIf LColorPattern(IRoll) = mdApp.LGreen Then
        mdApp.SetBlackRollBoard mdApp.INone
        mdApp.SetBlackPattern mdApp.INone
        
        mdApp.SetBlackBox WIN
        mdApp.SetRedBox WIN
    End If
    
    If IRoll = 0 Then
        mdApp.SetEvenRollBoard mdApp.INone
        mdApp.SetEvenPattern mdApp.INone
        
        mdApp.SetEvenBox WIN
        mdApp.SetOddBox WIN
    ElseIf IRoll Mod 2 = 0 Then
        mdApp.SetEvenRollBoard mdApp.ITick
        mdApp.SetEvenPattern mdApp.ITick
        
        mdApp.SetEvenBox WIN
        mdApp.SetOddBox LOSS
    Else
        mdApp.SetEvenRollBoard mdApp.ICross
        mdApp.SetEvenPattern mdApp.ICross
        
        mdApp.SetEvenBox LOSS
        mdApp.SetOddBox WIN
    End If
    
    If IRoll = 0 Then
        mdApp.SetHighRollBoard mdApp.INone
        mdApp.SetHighPattern mdApp.INone
        
        mdApp.SetHighBox WIN
        mdApp.SetLowBox WIN
    ElseIf IRoll >= 19 Then
        mdApp.SetHighRollBoard mdApp.ITick
        mdApp.SetHighPattern mdApp.ITick
        
        mdApp.SetHighBox WIN
        mdApp.SetLowBox LOSS
    Else
        mdApp.SetHighRollBoard mdApp.ICross
        mdApp.SetHighPattern mdApp.ICross
        
        mdApp.SetHighBox LOSS
        mdApp.SetLowBox WIN
    End If
    
    Dim SText() As String
    
    ReDim Preserve SText(0) As String
    
    If (mdApp.GetBlackBox = mdApp.IBoxWin) And (mdApp.GetRedBox(10) = 0) Then
        If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
        
        SText(UBound(SText)) = mdApp.SRedText
        
        mdApp.IRedBox(10) = mdApp.ITick
        
        If BAlert Then
            If Not BRed And BEngage Then
                fRed.Show vbModeless, Me
                
                BRed = True
            End If
        End If
    End If
    
    If (mdApp.GetRedBox = mdApp.IBoxWin) And (mdApp.GetBlackBox(10) = 0) Then
        If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
        
        SText(UBound(SText)) = mdApp.SBlackText
        
        mdApp.IBlackBox(10) = mdApp.ITick
        
        If BAlert Then
            If Not BBlack And BEngage Then
                fBlack.Show vbModeless, Me
                
                BBlack = True
            End If
        End If
    End If
    
    If (mdApp.GetEvenBox = mdApp.IBoxWin) And (mdApp.GetOddBox(10) = 0) Then
        If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
        
        SText(UBound(SText)) = mdApp.SOddText
        
        mdApp.IOddBox(10) = mdApp.ITick
        
        If BAlert Then
            If Not BOdd And BEngage Then
                fOdd.Show vbModeless, Me
                
                BOdd = True
            End If
        End If
    End If
    
    If (mdApp.GetOddBox = mdApp.IBoxWin) And (mdApp.GetEvenBox(10) = 0) Then
        If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
        
        SText(UBound(SText)) = mdApp.SEvenText
        
        mdApp.IEvenBox(10) = mdApp.ITick
        
        If BAlert Then
            If Not BEven And BEngage Then
                fEven.Show vbModeless, Me
                
                BEven = True
            End If
        End If
    End If
    
    If (mdApp.GetHighBox = mdApp.IBoxWin) And (mdApp.GetLowBox(10) = 0) Then
        If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
        
        SText(UBound(SText)) = mdApp.SLowText
        
        mdApp.ILowBox(10) = mdApp.ITick
        
        If BAlert Then
            If Not BLow And BEngage Then
                fLow.Show vbModeless, Me
                
                BLow = True
            End If
        End If
    End If
    
    If (mdApp.GetLowBox = mdApp.IBoxWin) And (mdApp.GetHighBox(10) = 0) Then
        If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
        
        SText(UBound(SText)) = mdApp.SHighText
        
        mdApp.IHighBox(10) = mdApp.ITick
        
        If BAlert Then
            If Not BHigh And BEngage Then
                fHigh.Show vbModeless, Me
                
                BHigh = True
            End If
        End If
    End If
    
    Dim ICounter As Integer
    
    Dim SValue As String
    
    If Not (Trim(SText(LBound(SText))) = "") Then
        If BAlert And BEngage Then
            SValue = ""
            
            For ICounter = LBound(SText) To UBound(SText)
                If Not (Trim(SValue) = "") Then SValue = _
                    SValue & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf
                
                SValue = SValue & " " & SText(ICounter)
            Next ICounter
            
            If UBound(SText) > 0 Then
                SValue = " BET" & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    SValue & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    vbCrLf & _
                    " FOREVER"
    
                fMessage.SetShape SValue, vbRed, 80, 5400, 5000, 0, 0
            Else
                SValue = " BET " & SValue & " FOREVER"
                
                fMessage.SetShape SValue, vbRed, 80, 12000, 2000, 0, 0
            End If
            
            fMessage.SetTimer 10, 300, 300, 300, False
            fMessage.Show vbModeless, Me
        End If
    Else
        ReDim SText(0) As String
        SText(0) = ""
        
        If (mdApp.GetBlackBox(5) = mdApp.IBoxStopBet) And (mdApp.GetBlackBox(11) = 0) Then
            If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
            
            SText(UBound(SText)) = mdApp.SBlackText
            
            mdApp.IBlackBox(11) = mdApp.ITick
            
            mdApp.IBlackBox(5) = 0
            mdApp.IBlackBox(6) = 0
            mdApp.IBlackBox(7) = 0
            mdApp.IBlackBox(8) = 0
            mdApp.IRedBox(0) = 0
            mdApp.IRedBox(1) = 0
            mdApp.IRedBox(2) = 0
            mdApp.IRedBox(3) = 0
            
            If mdApp.CheckTableHistory = 1 Then
                mdApp.IRollWBHT1(0) = mdApp.IRollWBHT1(0) + 1
            ElseIf mdApp.CheckTableHistory = 2 Then
                mdApp.IRollWBHT2(0) = mdApp.IRollWBHT2(0) + 1
            End If
            
            If BAlert Then
                If Not BBlack And BEngage Then
                    fBlack.Show vbModeless, Me
                    
                    BBlack = True
                End If
            End If
        End If
        
        If (mdApp.GetRedBox(5) = mdApp.IBoxStopBet) And (mdApp.GetRedBox(11) = 0) Then
            If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
            
            SText(UBound(SText)) = mdApp.SRedText
            
            mdApp.IRedBox(11) = mdApp.ITick
            
            mdApp.IRedBox(5) = 0
            mdApp.IRedBox(6) = 0
            mdApp.IRedBox(7) = 0
            mdApp.IRedBox(8) = 0
            mdApp.IBlackBox(0) = 0
            mdApp.IBlackBox(1) = 0
            mdApp.IBlackBox(2) = 0
            mdApp.IBlackBox(3) = 0
            
            If mdApp.CheckTableHistory = 1 Then
                mdApp.IRollWBHT1(1) = mdApp.IRollWBHT1(1) + 1
            ElseIf mdApp.CheckTableHistory = 2 Then
                mdApp.IRollWBHT2(1) = mdApp.IRollWBHT2(1) + 1
            End If
            
            If BAlert Then
                If Not BRed And BEngage Then
                    fRed.Show vbModeless, Me
                    
                    BRed = True
                End If
            End If
        End If
        
        If (mdApp.GetEvenBox(5) = mdApp.IBoxStopBet) And (mdApp.GetEvenBox(11) = 0) Then
            If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
            
            SText(UBound(SText)) = mdApp.SEvenText
            
            mdApp.IEvenBox(11) = mdApp.ITick
            
            mdApp.IEvenBox(5) = 0
            mdApp.IEvenBox(6) = 0
            mdApp.IEvenBox(7) = 0
            mdApp.IEvenBox(8) = 0
            mdApp.IOddBox(0) = 0
            mdApp.IOddBox(1) = 0
            mdApp.IOddBox(2) = 0
            mdApp.IOddBox(3) = 0
            
            If mdApp.CheckTableHistory = 1 Then
                mdApp.IRollWBHT1(2) = mdApp.IRollWBHT1(2) + 1
            ElseIf mdApp.CheckTableHistory = 2 Then
                mdApp.IRollWBHT2(2) = mdApp.IRollWBHT2(2) + 1
            End If
            
            If BAlert Then
                If Not BEven And BEngage Then
                    fEven.Show vbModeless, Me
                    
                    BEven = True
                End If
            End If
        End If
        
        If (mdApp.GetOddBox(5) = mdApp.IBoxStopBet) And (mdApp.GetOddBox(11) = 0) Then
            If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
            
            SText(UBound(SText)) = mdApp.SOddText
            
            mdApp.IOddBox(11) = mdApp.ITick
            
            mdApp.IOddBox(5) = 0
            mdApp.IOddBox(6) = 0
            mdApp.IOddBox(7) = 0
            mdApp.IOddBox(8) = 0
            mdApp.IEvenBox(0) = 0
            mdApp.IEvenBox(1) = 0
            mdApp.IEvenBox(2) = 0
            mdApp.IEvenBox(3) = 0
            
            If mdApp.CheckTableHistory = 1 Then
                mdApp.IRollWBHT1(3) = mdApp.IRollWBHT1(3) + 1
            ElseIf mdApp.CheckTableHistory = 2 Then
                mdApp.IRollWBHT2(3) = mdApp.IRollWBHT2(3) + 1
            End If
            
            If BAlert Then
                If Not BOdd And BEngage Then
                    fOdd.Show vbModeless, Me
                    
                    BOdd = True
                End If
            End If
        End If
        
        If (mdApp.GetHighBox(5) = mdApp.IBoxStopBet) And (mdApp.GetHighBox(11) = 0) Then
            If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
            
            SText(UBound(SText)) = mdApp.SHighText
            
            mdApp.IHighBox(11) = mdApp.ITick
            
            mdApp.IHighBox(5) = 0
            mdApp.IHighBox(6) = 0
            mdApp.IHighBox(7) = 0
            mdApp.IHighBox(8) = 0
            mdApp.ILowBox(0) = 0
            mdApp.ILowBox(1) = 0
            mdApp.ILowBox(2) = 0
            mdApp.ILowBox(3) = 0
            
            If mdApp.CheckTableHistory = 1 Then
                mdApp.IRollWBHT1(4) = mdApp.IRollWBHT1(4) + 1
            ElseIf mdApp.CheckTableHistory = 2 Then
                mdApp.IRollWBHT2(4) = mdApp.IRollWBHT2(4) + 1
            End If
            
            If BAlert Then
                If Not BHigh And BEngage Then
                    fHigh.Show vbModeless, Me
                    
                    BHigh = True
                End If
            End If
        End If
        
        If (mdApp.GetLowBox(5) = mdApp.IBoxStopBet) And (mdApp.GetLowBox(11) = 0) Then
            If Not (SText(UBound(SText)) = "") Then ReDim Preserve SText(UBound(SText) + 1) As String
            
            SText(UBound(SText)) = mdApp.SLowText
            
            mdApp.ILowBox(11) = mdApp.ITick
            
            mdApp.ILowBox(5) = 0
            mdApp.ILowBox(6) = 0
            mdApp.ILowBox(7) = 0
            mdApp.ILowBox(8) = 0
            mdApp.IHighBox(0) = 0
            mdApp.IHighBox(1) = 0
            mdApp.IHighBox(2) = 0
            mdApp.IHighBox(3) = 0
            
            If mdApp.CheckTableHistory = 1 Then
                mdApp.IRollWBHT1(5) = mdApp.IRollWBHT1(5) + 1
            ElseIf mdApp.CheckTableHistory = 2 Then
                mdApp.IRollWBHT2(5) = mdApp.IRollWBHT2(5) + 1
            End If
            
            If BAlert Then
                If Not BLow And BEngage Then
                    fLow.Show vbModeless, Me
                    
                    BLow = True
                End If
            End If
        End If
        
        If Not (Trim(SText(LBound(SText))) = "") Then
            If BAlert And BEngage Then
                SValue = ""
                
                For ICounter = LBound(SText) To UBound(SText)
                    If Not (Trim(SValue) = "") Then SValue = _
                        SValue & _
                        vbCrLf & _
                        vbCrLf & _
                        vbCrLf & _
                        vbCrLf & _
                        vbCrLf
                    
                    SValue = SValue & " " & SText(ICounter)
                Next ICounter
                
                If UBound(SText) > 0 Then
                    SValue = " STOP BETS ON" & _
                        vbCrLf & _
                        vbCrLf & _
                        vbCrLf & _
                        vbCrLf & _
                        vbCrLf & _
                        SValue
        
                    fMessage.SetShape SValue, vbRed, 80, 5400, 5000, 0, 0
                Else
                    SValue = " STOP BET ON " & SValue
                    
                    fMessage.SetShape SValue, vbRed, 80, 12000, 2000, 0, 0
                End If
                
                fMessage.SetTimer 10, 300, 300, 300, False
                fMessage.Show vbModeless, Me
            End If
        End If
    End If
    
    If BAlert Then
        If BBlack Then fBlack.SetText
        If BRed Then fRed.SetText
        If BEven Then fEven.SetText
        If BOdd Then fOdd.SetText
        If BHigh Then fHigh.SetText
        If BLow Then fLow.SetText
    End If
End Sub

Private Sub SetRollBoard()
    Dim ICounter As Integer
    Dim ICDraw As Integer
    Dim IRDraw As Integer
    
    If UBound(mdApp.INumberRollBoard) >= mdApp.IRowMax Then
        IRDraw = (UBound(mdApp.INumberRollBoard) - mdApp.IRowMax) + 1
    Else
        IRDraw = LBound(mdApp.INumberRollBoard)
    End If
    
    ICDraw = 0
    
    If mdApp.INumberRollBoard(LBound(mdApp.INumberRollBoard)) = mdApp.IBlank Then
    Else
        For ICounter = IRDraw To UBound(mdApp.INumberRollBoard)
            Me.lbRoll(ICDraw).ForeColor = mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter))
            Me.lbRoll(ICDraw).Caption = CStr(mdApp.INumberRollBoard(ICounter))
    
            If mdApp.INumberRollBoard(ICounter) = mdApp.INone Then
                If mdApp.BBlackLastFocus(ICounter) Then
                    Set Me.pcBlack(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcBlack(ICDraw).Picture = Me.pcNone
                End If
                
                If mdApp.BEvenLastFocus(ICounter) Then
                    Set Me.pcEven(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcEven(ICDraw).Picture = Me.pcNone
                End If
                
                If mdApp.BHighLastFocus(ICounter) Then
                    Set Me.pcHigh(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcHigh(ICDraw).Picture = Me.pcNone
                End If
            Else
                If mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LRed Then
                    If mdApp.BBlackLastFocus(ICounter) Then
                        Set Me.pcBlack(ICDraw).Picture = Me.pcCrossLast
                    Else
                        Set Me.pcBlack(ICDraw).Picture = Me.pcCross
                    End If
                Else
                    If mdApp.BBlackLastFocus(ICounter) Then
                        Set Me.pcBlack(ICDraw).Picture = Me.pcTickLast
                    Else
                        Set Me.pcBlack(ICDraw).Picture = Me.pcTick
                    End If
                End If
    
                If mdApp.INumberRollBoard(ICounter) Mod 2 = 0 Then
                    If mdApp.BEvenLastFocus(ICounter) Then
                        Set Me.pcEven(ICDraw).Picture = Me.pcTickLast
                    Else
                        Set Me.pcEven(ICDraw).Picture = Me.pcTick
                    End If
                Else
                    If mdApp.BEvenLastFocus(ICounter) Then
                        Set Me.pcEven(ICDraw).Picture = Me.pcCrossLast
                    Else
                        Set Me.pcEven(ICDraw).Picture = Me.pcCross
                    End If
                End If
    
                If mdApp.INumberRollBoard(ICounter) >= 19 Then
                    If mdApp.BHighLastFocus(ICounter) Then
                        Set Me.pcHigh(ICDraw).Picture = Me.pcTickLast
                    Else
                        Set Me.pcHigh(ICDraw).Picture = Me.pcTick
                    End If
                Else
                    If mdApp.BHighLastFocus(ICounter) Then
                        Set Me.pcHigh(ICDraw).Picture = Me.pcCrossLast
                    Else
                        Set Me.pcHigh(ICDraw).Picture = Me.pcCross
                    End If
                End If
            End If
            
            ICDraw = ICDraw + 1
        Next ICounter
    End If
End Sub

Private Sub SetRollFocus()
    Dim IBlackFocus As Integer
    Dim IEvenFocus As Integer
    Dim IHighFocus As Integer
    Dim ICounter As Integer
    Dim ICDraw As Integer

    IBlackFocus = mdApp.CheckFocus(IBlackType)
    If IBlackFocus = 1 Then
        IBlackFocus = 5
    ElseIf IBlackFocus > 1 Then
        If IBlackFocus >= (mdApp.IRowMax - 5) Then
            IBlackFocus = (mdApp.IRowMax - 1)
        Else
            IBlackFocus = 5 + (IBlackFocus - 1)
        End If
    End If

    IEvenFocus = mdApp.CheckFocus(IEvenType)
    If IEvenFocus = 1 Then
        IEvenFocus = 5
    ElseIf IEvenFocus > 1 Then
        If IEvenFocus >= (mdApp.IRowMax - 5) Then
            IEvenFocus = (mdApp.IRowMax - 1)
        Else
            IEvenFocus = 5 + (IEvenFocus - 1)
        End If
    End If

    IHighFocus = mdApp.CheckFocus(IHighType)
    If IHighFocus = 1 Then
        IHighFocus = 5
    ElseIf IHighFocus > 1 Then
        If IHighFocus >= (mdApp.IRowMax - 5) Then
            IHighFocus = (mdApp.IRowMax - 1)
        Else
            IHighFocus = 5 + (IHighFocus - 1)
        End If
    End If
    
    If mdApp.CheckBlack Then
        If (UBound(mdApp.INumberRollBoard) - IBlackFocus) >= 0 Then
            If BHistory Then fHistory.FillPattern UBound(mdApp.INumberRollBoard) - IBlackFocus, mdApp.IBlackType
            
            If UBound(mdApp.INumberRollBoard) >= mdApp.IRowMax Then
                ICDraw = mdApp.IRowMax - 1
            Else
                ICDraw = UBound(mdApp.INumberRollBoard)
            End If
            
            For ICounter = UBound(mdApp.INumberRollBoard) To (UBound(mdApp.INumberRollBoard) - IBlackFocus) Step -1
                mdApp.BBlackLastFocus(ICounter) = True
                
                If mdApp.INumberRollBoard(ICounter) = mdApp.INone Then
                    mdApp.IRollBPatternHTS(ICounter) = INone
                    
                    Set Me.pcBlack(ICDraw).Picture = Me.pcNoneFocus
                Else
                    If mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LGreen Then
                        mdApp.IRollBPatternHTS(ICounter) = INone
                        
                        Set Me.pcBlack(ICDraw).Picture = Me.pcNoneFocus
                    ElseIf mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LRed Then
                        mdApp.IRollBPatternHTS(ICounter) = ICross
                        
                        Set Me.pcBlack(ICDraw).Picture = Me.pcCrossFocus
                    ElseIf mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LBlack Then
                        mdApp.IRollBPatternHTS(ICounter) = ITick
                        
                        Set Me.pcBlack(ICDraw).Picture = Me.pcTickFocus
                    End If
                End If
                
                ICDraw = ICDraw - 1
            Next ICounter
        End If
    End If

    If mdApp.CheckEven Then
        If (UBound(mdApp.INumberRollBoard) - IEvenFocus) >= 0 Then
            If BHistory Then fHistory.FillPattern UBound(mdApp.INumberRollBoard) - IEvenFocus, mdApp.IEvenType
            
            If UBound(mdApp.INumberRollBoard) >= mdApp.IRowMax Then
                ICDraw = mdApp.IRowMax - 1
            Else
                ICDraw = UBound(mdApp.INumberRollBoard)
            End If
            
            For ICounter = UBound(mdApp.INumberRollBoard) To (UBound(mdApp.INumberRollBoard) - IEvenFocus) Step -1
                mdApp.BEvenLastFocus(ICounter) = True
                
                If mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LGreen Then
                    mdApp.IRollEPatternHTS(ICounter) = INone
                    
                    Set Me.pcEven(ICDraw).Picture = Me.pcNoneFocus
                ElseIf mdApp.INumberRollBoard(ICounter) Mod 2 = 0 Then
                    mdApp.IRollEPatternHTS(ICounter) = ITick
                    
                    Set Me.pcEven(ICDraw).Picture = Me.pcTickFocus
                Else
                    mdApp.IRollEPatternHTS(ICounter) = ICross
                    
                    Set Me.pcEven(ICDraw).Picture = Me.pcCrossFocus
                End If
                
                ICDraw = ICDraw - 1
            Next ICounter
        End If
    End If
    
    If mdApp.CheckHigh Then
        If (UBound(mdApp.INumberRollBoard) - IEvenFocus) >= 0 Then
            If BHistory Then fHistory.FillPattern UBound(mdApp.INumberRollBoard) - IHighFocus, mdApp.IHighType
            
            If UBound(mdApp.INumberRollBoard) >= mdApp.IRowMax Then
                ICDraw = mdApp.IRowMax - 1
            Else
                ICDraw = UBound(mdApp.INumberRollBoard)
            End If
            
            For ICounter = UBound(mdApp.INumberRollBoard) To (UBound(mdApp.INumberRollBoard) - IHighFocus) Step -1
                mdApp.BHighLastFocus(ICounter) = True
                
                If mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LGreen Then
                    mdApp.IRollHPatternHTS(ICounter) = INone
                    
                    Set Me.pcHigh(ICDraw).Picture = Me.pcNoneFocus
                ElseIf mdApp.INumberRollBoard(ICounter) >= 19 Then
                    mdApp.IRollHPatternHTS(ICounter) = ITick
                    
                    Set Me.pcHigh(ICDraw).Picture = Me.pcTickFocus
                Else
                    mdApp.IRollHPatternHTS(ICounter) = ICross
                    
                    Set Me.pcHigh(ICDraw).Picture = Me.pcCrossFocus
                End If
                
                ICDraw = ICDraw - 1
            Next ICounter
        End If
    End If
    
    mdApp.SavePatternHistory
End Sub

Private Sub SetRollHistory()
    Dim ICounter As Integer
    
    If mdApp.CheckTableHistory = 1 And BTable1 Then
        For ICounter = LBound(mdApp.IRollHT1) To UBound(mdApp.IRollHT1)
            If mdApp.IRollHT1(ICounter) > 0 Then
                fRoll1.lbRollCounter(ICounter).Caption = Format(mdApp.IRollHT1(ICounter), "#,##0")
            Else
                fRoll1.lbRollCounter(ICounter).Caption = ""
            End If
        Next ICounter
        
        For ICounter = LBound(mdApp.IRollCHT1) To UBound(mdApp.IRollCHT1)
            fRoll1.lbCounter(ICounter).Caption = " " & Format(mdApp.IRollCHT1(ICounter), "#,##0")
        Next ICounter
        
        For ICounter = LBound(mdApp.IRollWBHT1) To UBound(mdApp.IRollWBHT1)
            fRoll1.lbWinBet(ICounter).Caption = " " & CStr(mdApp.IRollWBHT1(ICounter))
        Next ICounter
        
        For ICounter = LBound(mdApp.IRollLBHT1) To UBound(mdApp.IRollLBHT1)
            fRoll1.lbLossBet(ICounter).Caption = " " & CStr(mdApp.IRollLBHT1(ICounter))
        Next ICounter

        fRoll1.lbNow.Caption = Format(CInt(fRoll1.lbNow.Caption) + 1, "#,##0")
    ElseIf mdApp.CheckTableHistory = 2 And BTable2 Then
        For ICounter = LBound(mdApp.IRollHT2) To UBound(mdApp.IRollHT2)
            If mdApp.IRollHT2(ICounter) > 0 Then
                fRoll2.lbRollCounter(ICounter).Caption = Format(mdApp.IRollHT2(ICounter), "#,##0")
            Else
                fRoll2.lbRollCounter(ICounter).Caption = ""
            End If
        Next ICounter
        
        For ICounter = LBound(mdApp.IRollCHT2) To UBound(mdApp.IRollCHT2)
            fRoll2.lbCounter(ICounter).Caption = " " & Format(mdApp.IRollCHT2(ICounter), "#,##0")
        Next ICounter
        
        For ICounter = LBound(mdApp.IRollWBHT2) To UBound(mdApp.IRollWBHT2)
            fRoll2.lbWinBet(ICounter).Caption = " " & CStr(mdApp.IRollWBHT2(ICounter))
        Next ICounter
        
        For ICounter = LBound(mdApp.IRollLBHT1) To UBound(mdApp.IRollLBHT1)
            fRoll2.lbLossBet(ICounter).Caption = " " & CStr(mdApp.IRollLBHT2(ICounter))
        Next ICounter

        fRoll2.lbNow.Caption = Format(CInt(fRoll2.lbNow.Caption) + 1, "#,##0")
    End If
End Sub

Private Sub SetBetText(Optional ByVal BClear As Boolean = True, Optional ByVal BAlert As Boolean = True)
    Dim ICounter As Integer
    
    If BAlert Then
        For ICounter = 0 To Me.lbBet.Count - 1
            Me.lbBet(ICounter).Caption = ""
        Next ICounter
    End If
    
    If BClear Then
    Else
        Dim SPattern() As String
        Dim BDoubleBet As Boolean
        
        SPattern = Split(mdApp.CheckPattern(BAlert), vbCrLf)
        
        If BAlert Then
            If Not (Trim(SPattern(0)) = "") Or Not (Trim(SPattern(1)) = "") Or Not (Trim(SPattern(2)) = "") Then
                If Not (Trim(SPattern(0)) = "") And Not (Trim(SPattern(1)) = "") Then
                    BDoubleBet = True
                ElseIf Not Trim(SPattern(1)) = "" And Not Trim(SPattern(2)) = "" Then
                    BDoubleBet = True
                ElseIf Not (Trim(SPattern(0)) = "") And Not (Trim(SPattern(2)) = "") Then
                    BDoubleBet = True
                Else
                    BDoubleBet = False
                End If
            End If
            
            For ICounter = 0 To Me.lbBet.Count - 5
                Me.lbBet(ICounter).Caption = SPattern(ICounter)
                
                If (Trim(SPattern(ICounter)) = "") And BDoubleBet Then
                    'Me.lbBet(ICounter).Caption = "0"
                End If
                
                If Not (Trim(SPattern(ICounter)) = "") Then
                    If (mdApp.CheckBOutcome) > 0 And (ICounter = 0) Then
                        Me.lbBet(ICounter + 4).Caption = CStr(mdApp.CheckBOutcome + 1)
                    ElseIf (mdApp.CheckEOutcome) > 0 And (ICounter = 1) Then
                        Me.lbBet(ICounter + 4).Caption = CStr(mdApp.CheckEOutcome + 1)
                    ElseIf (mdApp.CheckHOutcome) > 0 And (ICounter = 2) Then
                        Me.lbBet(ICounter + 4).Caption = CStr(mdApp.CheckHOutcome + 1)
                    End If
                End If
            Next ICounter
            
            If Not (Trim(SPattern(0)) = "") And Not (Trim(SPattern(1)) = "") And Not (Trim(SPattern(2)) = "") Then
                Me.lbBet(3).Visible = True
                'Me.lbBet(3).Caption = "0"
            End If
        End If
    End If
End Sub

Private Sub SetClear()
    If mdSecurity.EncryptText(Me.txPassword.Text, mdSecurity.SKey) = mdSecurity.SSecured Then
        Dim ICounter As Integer
        
        For ICounter = 0 To mdApp.IRowMax - 1
            Me.lbRoll(ICounter).Caption = ""
            
            Set Me.pcBlack(ICounter).Picture = LoadPicture
            Set Me.pcEven(ICounter).Picture = LoadPicture
            Set Me.pcHigh(ICounter).Picture = LoadPicture
        Next ICounter
        
        mdApp.InitRollBoard
        mdApp.InitRollPattern
        
        SetBetText
        
        mdApp.ClearHistory
        mdApp.InitPatternHistory
        SetRollFocus
        mdApp.SetRefreshWinLoss
        mdApp.ClearBoardHistory
        
        mdApp.ICounterRollDec = mdApp.IRollBaseDec
        
        Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard), "#,##0")
        Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
        
        If BBox Then
            fBox.lbWin.Caption = CStr(mdApp.CheckWin)
            fBox.lbLoss.Caption = CStr(mdApp.CheckLoss)
        End If
        
        If BHistory Then fHistory.lvHistory.ListItems.Clear
        
        ClearBox
    End If
    
    Me.txPassword.Visible = False
End Sub

Private Sub SetDelete()
    Dim IBlackRollTemp As Integer
    Dim IEvenRollTemp As Integer
    Dim IHighRollTemp As Integer
    
    If BHistory Then
        fHistory.DeleteItem
    Else
        mdApp.DeleteHistory
    End If
    
    If UBound(mdApp.INumberRollBoard) <= 0 Then
        mdApp.INumberRollBoard(0) = mdApp.IBlank
        mdApp.IBlackRollBoard(0) = mdApp.IBlank
        mdApp.IEvenRollBoard(0) = mdApp.IBlank
        mdApp.IHighRollBoard(0) = mdApp.IBlank
        mdApp.IBlackRollPattern(0) = mdApp.IBlank
        mdApp.IEvenRollPattern(0) = mdApp.IBlank
        mdApp.IHighRollPattern(0) = mdApp.IBlank
        
        Me.lbRoll(0).Caption = ""
        
        Set Me.pcBlack(0).Picture = LoadPicture
        Set Me.pcEven(0).Picture = LoadPicture
        Set Me.pcHigh(0).Picture = LoadPicture
        
        SetBetText
        ClearBox
        
        mdApp.BBlackLastFocus(0) = False
        mdApp.BEvenLastFocus(0) = False
        mdApp.BHighLastFocus(0) = False
    Else
        IBlackRollTemp = mdApp.IBlackRollBoard(UBound(mdApp.IBlackRollBoard))
        IEvenRollTemp = mdApp.IEvenRollBoard(UBound(mdApp.IEvenRollBoard))
        IHighRollTemp = mdApp.IHighRollBoard(UBound(mdApp.IHighRollBoard))
        
        mdApp.BBlackLastFocus(UBound(mdApp.INumberRollBoard)) = False
        mdApp.BEvenLastFocus(UBound(mdApp.INumberRollBoard)) = False
        mdApp.BHighLastFocus(UBound(mdApp.INumberRollBoard)) = False
        
        ReDim Preserve mdApp.INumberRollBoard(UBound(mdApp.INumberRollBoard) - 1) As Integer
        ReDim Preserve mdApp.IBlackRollBoard(UBound(mdApp.IBlackRollBoard) - 1) As Integer
        ReDim Preserve mdApp.IEvenRollBoard(UBound(mdApp.IEvenRollBoard) - 1) As Integer
        ReDim Preserve mdApp.IHighRollBoard(UBound(mdApp.IHighRollBoard) - 1) As Integer
        ReDim Preserve mdApp.IBlackRollPattern(UBound(mdApp.IBlackRollPattern) - 1) As Integer
        ReDim Preserve mdApp.IEvenRollPattern(UBound(mdApp.IEvenRollPattern) - 1) As Integer
        ReDim Preserve mdApp.IHighRollPattern(UBound(mdApp.IHighRollPattern) - 1) As Integer
        ReDim Preserve mdApp.IRollBPatternHTS(UBound(mdApp.INumberRollBoard)) As Integer
        ReDim Preserve mdApp.IRollEPatternHTS(UBound(mdApp.INumberRollBoard)) As Integer
        ReDim Preserve mdApp.IRollHPatternHTS(UBound(mdApp.INumberRollBoard)) As Integer
        
        If UBound(mdApp.INumberRollBoard) >= (mdApp.IRowMax) Then
            SetRollBoard
        Else
            Dim IRCounter As Integer
            Dim ICounter As Integer
            
            IRCounter = (mdApp.IRowMax - 1) - (mdApp.IRowMax - UBound(mdApp.INumberRollBoard) - 1)
            
            For ICounter = 0 To mdApp.IRowMax - 1
                If ICounter > IRCounter Then
                    Me.lbRoll(ICounter).Caption = ""
                    
                    Set Me.pcBlack(ICounter).Picture = LoadPicture
                    Set Me.pcEven(ICounter).Picture = LoadPicture
                    Set Me.pcHigh(ICounter).Picture = LoadPicture
                Else
                    Me.lbRoll(ICounter).ForeColor = mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter))
                    Me.lbRoll(ICounter).Caption = CStr(mdApp.INumberRollBoard(ICounter))
            
                    If mdApp.INumberRollBoard(ICounter) = INone Then
                        Set Me.pcBlack(ICounter).Picture = Me.pcNone
                        Set Me.pcEven(ICounter).Picture = Me.pcNone
                        Set Me.pcHigh(ICounter).Picture = Me.pcNone
                    Else
                        If mdApp.LColorPattern(mdApp.INumberRollBoard(ICounter)) = mdApp.LRed Then
                            Set Me.pcBlack(ICounter).Picture = Me.pcCross
                        Else
                            Set Me.pcBlack(ICounter).Picture = Me.pcTick
                        End If
                        
                        If mdApp.INumberRollBoard(ICounter) Mod 2 = 0 Then
                            Set Me.pcEven(ICounter).Picture = Me.pcTick
                        Else
                            Set Me.pcEven(ICounter).Picture = Me.pcCross
                        End If
                        
                        If mdApp.INumberRollBoard(ICounter) >= 19 Then
                            Set Me.pcHigh(ICounter).Picture = Me.pcTick
                        Else
                            Set Me.pcHigh(ICounter).Picture = Me.pcCross
                        End If
                    End If
                End If
            Next ICounter
        End If
        
        SetBetText False
        
        mdApp.CheckSubtWinLoss IBlackRollTemp, IEvenRollTemp, IHighRollTemp
        
        If IBlackRollTemp = mdApp.ITick Then
            mdApp.SetRBlackBox WIN
            mdApp.SetRRedBox LOSS
        ElseIf IBlackRollTemp = mdApp.INone Then
            mdApp.SetRBlackBox WIN
            mdApp.SetRRedBox WIN
        Else
            mdApp.SetRBlackBox LOSS
            mdApp.SetRRedBox WIN
        End If
        
        If BBlack Then fBlack.SetText
        If BRed Then fRed.SetText
        
        If IEvenRollTemp = mdApp.ITick Then
            mdApp.SetREvenBox WIN
            mdApp.SetROddBox LOSS
        ElseIf IEvenRollTemp = mdApp.INone Then
            mdApp.SetREvenBox WIN
            mdApp.SetROddBox WIN
        Else
            mdApp.SetREvenBox LOSS
            mdApp.SetROddBox WIN
        End If
        
        If BEven Then fEven.SetText
        If BOdd Then fOdd.SetText
        
        If IHighRollTemp = mdApp.ITick Then
            mdApp.SetRHighBox WIN
        ElseIf IEvenRollTemp = mdApp.INone Then
            mdApp.SetRHighBox WIN
            mdApp.SetRLowBox WIN
        Else
            mdApp.SetRHighBox LOSS
            mdApp.SetRLowBox WIN
        End If
        
        If BHigh Then fHigh.SetText
        If BLow Then fLow.SetText
    End If
    
    If mdApp.INumberRollBoard(UBound(mdApp.INumberRollBoard)) = IBlank Then
        Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard), "#,##0")
        
        mdApp.ICounterRollDec = mdApp.IRollBaseDec
        
        mdApp.SetRefreshWinLoss
    Else
        Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard) + 1, "#,##0")
        
        If mdApp.ICounterRollDec < mdApp.IRollBaseDec Then mdApp.ICounterRollDec = mdApp.ICounterRollDec + 1
    End If
    
    Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
    
    mdApp.DeleteRollHistory
    SetRollFocus
    mdApp.SaveBoxHistory
    
    If BBox Then
        fBox.lbWin.Caption = CStr(mdApp.CheckWin)
        fBox.lbLoss.Caption = CStr(mdApp.CheckLoss)
    End If
End Sub

Private Sub ChangePassword()
    If Not mdSecurity.GetGuest Then fCPassword.Show vbModal
End Sub

Private Sub ShowAfterNoAlert()
    If BBlack Then fBlack.SetText
    If BRed Then fRed.SetText
    If BEven Then fEven.SetText
    If BOdd Then fOdd.SetText
    If BHigh Then fHigh.SetText
    If BLow Then fLow.SetText
    
    SetBetText False
    SetRollBoard
    SetRollFocus
    SetRollHistory
    
    Me.lbCount(0).Caption = " " & Format(UBound(mdApp.INumberRollBoard) + 1, "#,##0")
    Me.lbCount(1).Caption = " " & Format(mdApp.ICounterRollDec, "#,##0")
End Sub
