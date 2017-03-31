VERSION 5.00
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "THE MATRIX"
   ClientHeight    =   8775
   ClientLeft      =   105
   ClientTop       =   -285
   ClientWidth     =   7575
   Icon            =   "fMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   7575
   Begin VB.Timer trDelay 
      Left            =   6960
      Top             =   840
   End
   Begin VB.Frame frBet 
      BackColor       =   &H00F0F0F0&
      Caption         =   "BET"
      Height          =   1455
      Left            =   5160
      TabIndex        =   141
      Top             =   7200
      Width           =   2295
      Begin VB.Timer trBlink 
         Index           =   2
         Left            =   1080
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   0
         Left            =   120
         Top             =   120
      End
      Begin VB.Timer trBlink 
         Index           =   1
         Left            =   600
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
         Index           =   0
         Left            =   120
         TabIndex        =   170
         Top             =   240
         Width           =   2055
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
         Left            =   120
         TabIndex        =   169
         Top             =   600
         Width           =   2055
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
         Left            =   120
         TabIndex        =   168
         Top             =   960
         Width           =   2055
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox pcGreen 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6960
      Picture         =   "fMain.frx":08CA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   140
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pcRed 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6960
      Picture         =   "fMain.frx":0EEB
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   139
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   138
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   137
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   136
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   135
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   134
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   133
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   132
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   131
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   130
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   129
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   128
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   127
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   126
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   125
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   124
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   123
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   122
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   121
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   120
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   119
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   118
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   117
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   116
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   115
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   114
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   113
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   112
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   111
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   110
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   109
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   108
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   107
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   106
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   105
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   104
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   103
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   840
      Width           =   375
   End
   Begin VB.PictureBox pcNoneFocus 
      Height          =   375
      Left            =   6480
      Picture         =   "fMain.frx":15E8
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcCrossFocus 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   5880
      Picture         =   "fMain.frx":18C5
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcTickFocus 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   5280
      Picture         =   "fMain.frx":1F05
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6360
      TabIndex        =   69
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox pcMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   6600
      Picture         =   "fMain.frx":23D4
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
      Left            =   7080
      Picture         =   "fMain.frx":2D46
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   65
      Top             =   120
      Width           =   420
   End
   Begin VB.PictureBox pcNone 
      Height          =   375
      Left            =   6480
      Picture         =   "fMain.frx":36B8
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
      Left            =   5880
      Picture         =   "fMain.frx":3859
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
      Left            =   5280
      Picture         =   "fMain.frx":3B4C
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
      Left            =   5280
      Picture         =   "fMain.frx":3D83
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcCrossLast 
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   5880
      Picture         =   "fMain.frx":3FA7
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pcNoneLast 
      Height          =   375
      Left            =   6480
      Picture         =   "fMain.frx":4222
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   144
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
      Width           =   4935
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   212
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   211
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   210
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   209
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   208
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   207
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   206
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   205
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   204
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   203
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   202
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   201
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   200
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   199
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   198
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   197
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   196
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   195
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   194
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcBlack6 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   4320
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   193
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   192
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   191
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   190
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   189
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   188
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   187
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   186
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   185
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   184
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   183
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   182
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   181
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   180
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   179
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   178
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   177
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   176
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   175
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   174
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcBlack5 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   3600
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   173
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   167
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   166
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   165
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   164
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   163
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   162
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   161
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   160
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   159
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   158
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   157
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   156
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   155
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   154
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   153
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   152
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   151
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   150
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   149
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack4 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   2880
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   148
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   101
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   100
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   99
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   98
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   97
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   96
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   95
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   94
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   93
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   92
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   91
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   90
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   19
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   89
         Top             =   7440
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   18
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   88
         Top             =   7080
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   17
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   87
         Top             =   6720
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   16
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   86
         Top             =   6360
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   15
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   85
         Top             =   6000
         Width           =   495
      End
      Begin VB.PictureBox pcBlack 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   14
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   84
         Top             =   5640
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   58
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   57
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   56
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   55
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   54
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   53
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   52
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   51
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   50
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   49
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   48
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   47
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   13
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   5280
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   12
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   4920
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   11
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   42
         Top             =   4560
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   10
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   4200
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   9
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   40
         Top             =   3840
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   8
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   39
         Top             =   3480
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   7
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   38
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   6
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   37
         Top             =   2760
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   5
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   36
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   35
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   1440
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   34
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   1440
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
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
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   1440
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
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.PictureBox pcBlack3 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   600
         Width           =   495
      End
      Begin VB.PictureBox pcBlack2 
         BackColor       =   &H00F0F0F0&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1440
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
         Left            =   720
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbMain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B6"
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
         Index           =   6
         Left            =   4320
         TabIndex        =   172
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbMain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B5"
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
         Index           =   5
         Left            =   3600
         TabIndex        =   171
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbMain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B4"
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
         Left            =   2880
         TabIndex        =   147
         Top             =   120
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B3"
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
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbMain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B2"
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
         Left            =   1440
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lbMain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B1"
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
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   495
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   525
      End
      Begin VB.Line lnBorder 
         Index           =   4
         X1              =   0
         X2              =   4920
         Y1              =   480
         Y2              =   480
      End
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
      Index           =   2
      Left            =   6360
      TabIndex        =   146
      Top             =   6120
      Width           =   975
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
      Left            =   5280
      TabIndex        =   145
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   83
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   15
      Index           =   4
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   8895
      Index           =   3
      Left            =   7560
      TabIndex        =   81
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   15
      Index           =   2
      Left            =   0
      TabIndex        =   80
      Top             =   8760
      Width           =   7575
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00404040&
      Height          =   8895
      Index           =   1
      Left            =   0
      TabIndex        =   79
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
      Left            =   5280
      TabIndex        =   68
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00FF8080&
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "THE MATRIX"
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
      Width           =   7575
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LStartButton As Long = &HFF00&
Private Const LStopButton As Long = &H8080FF
Private Const SStartButton As String = "Start"
Private Const SStopButton As String = "Stop"

Private SnX As Single
Private SnY As Single
Private BMove As Boolean

Private IBlackTimer As Integer

Public Sub CloseForm()
    Unload Me
End Sub

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BMove = True
    
    SnX = X
    SnY = Y
    
    Me.MousePointer = MousePointerConstants.vbSizeAll
End Sub

Private Sub lbBet_Change(Index As Integer)
    If Not (Trim(Me.lbBet(Index).Caption) = "") Then
        If Trim(Me.lbBet(Index).Caption) = "0" Then
            Me.lbBet(Index).ForeColor = LGreen
        Else
            Me.lbBet(Index).ForeColor = LBlack
        End If
        
        If (Index = 4) Or (Index = 5) Or (Index = 6) Then
            Me.lbBet(Index).ForeColor = LBlue
        End If
        
        Me.trBlink(Index).Enabled = True
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
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fMain = Nothing
End Sub

Private Sub pcExit_Click()
    Unload Me
End Sub

Private Sub pcMin_Click()
    Me.WindowState = FormWindowStateConstants.vbMinimized
End Sub

Private Sub cdRoll_Click(Index As Integer)
    SetRoll Index
End Sub

Private Sub cdDelete_Click()
    SetDelete
End Sub

Private Sub SetInitial()
    Dim ICounter As Integer
    
    mdGeneral.CenterWindows Me, False

    Init
    
    If UBound(INumberRollBoard) = 0 Then
        If INumberRollBoard(0) = IBlank Then
            Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard), "#,##0")
        Else
            Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard) + 1, "#,##0")
        End If
    Else
        Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard) + 1, "#,##0")
    End If
    Me.lbCount(1).Caption = "$ " & Format(ICounterRollDec, "#,##0")
    
    For ICounter = 0 To 2
        Me.trBlink(ICounter).Interval = 250
        Me.trBlink(ICounter).Enabled = False
    Next ICounter
    
    IBlackTimer = 250
End Sub

Private Sub SetRoll(ByVal IRoll As Integer, Optional ByVal BAlert As Boolean = True)
    If ICounterRollDec = 0 Then Exit Sub
    
    ICounterRollDec = ICounterRollDec - 1
    
    If IRoll > INumberMax Then Exit Sub
    
    SetBoardPattern IRoll, BAlert
    
    SetBetText False, BAlert
    
    If BAlert Then
        SetRollBoard
        SetRollFocus
    End If
    
    mdAPI.Beep 300, 50
    
    Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard) + 1, "#,##0")
    Me.lbCount(1).Caption = "$ " & Format(ICounterRollDec, "#,##0")
End Sub

Private Sub SetRollFocus()
    Dim ICounter As Integer
    Dim ICDraw As Integer
    
    If BBlack1 Then
        If (UBound(INumberRollBoard) - IBlack1Focus) >= 0 Then
            If UBound(INumberRollBoard) >= IRowMax Then
                ICDraw = IRowMax - 1
            Else
                ICDraw = UBound(INumberRollBoard)
            End If
            
            For ICounter = IBlack1Focus To UBound(INumberRollBoard)
                BBlack1LastFocus(ICounter) = True
                
                If INumberRollBoard(ICounter) = INone Then
                    IRollB1Pattern(ICounter) = INone
                    
                    Set Me.pcBlack(ICounter).Picture = Me.pcNoneFocus
                Else
                    If LColorPattern(INumberRollBoard(ICounter)) = LGreen Then
                        IRollB1Pattern(ICounter) = INone
                        
                        Set Me.pcBlack(ICounter).Picture = Me.pcNoneFocus
                    ElseIf LColorPattern(INumberRollBoard(ICounter)) = LRed Then
                        IRollB1Pattern(ICounter) = ICross
                        
                        Set Me.pcBlack(ICounter).Picture = Me.pcCrossFocus
                    ElseIf LColorPattern(INumberRollBoard(ICounter)) = LBlack Then
                        IRollB1Pattern(ICounter) = ITick
                        
                        Set Me.pcBlack(ICounter).Picture = Me.pcTickFocus
                    End If
                End If
            Next ICounter
        End If
    End If
End Sub

Private Sub SetRollBoard()
    Dim ICounter As Integer
    Dim ICDraw As Integer
    Dim IRDraw As Integer
    
    If UBound(INumberRollBoard) >= IRowMax Then
        IRDraw = (UBound(INumberRollBoard) - IRowMax) + 1
    Else
        IRDraw = LBound(INumberRollBoard)
    End If
    
    ICDraw = 0
    
    If INumberRollBoard(LBound(INumberRollBoard)) = IBlank Then
    Else
        For ICounter = IRDraw To UBound(INumberRollBoard)
            Me.lbRoll(ICDraw).ForeColor = LColorPattern(INumberRollBoard(ICounter))
            Me.lbRoll(ICDraw).Caption = CStr(INumberRollBoard(ICounter))
            
            If INumberRollBoard(ICounter) = INone Then
                If BBlack1LastFocus(ICounter) Then
                    Set Me.pcBlack(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcBlack(ICDraw).Picture = Me.pcNone
                End If
                
                If BBlack2LastFocus(ICounter) Then
                    Set Me.pcBlack2(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcBlack2(ICDraw).Picture = Me.pcNone
                End If
                
                If BBlack3LastFocus(ICounter) Then
                    Set Me.pcBlack3(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcBlack3(ICDraw).Picture = Me.pcNone
                End If
                
                If BBlack4LastFocus(ICounter) Then
                    Set Me.pcBlack4(ICDraw).Picture = Me.pcNoneLast
                Else
                    Set Me.pcBlack4(ICDraw).Picture = Me.pcNone
                End If
            Else
                If LColorPattern(INumberRollBoard(ICounter)) = LRed Then
                    If BBlack1LastFocus(ICounter) Then
                        Set Me.pcBlack(ICDraw).Picture = Me.pcCrossLast
                    Else
                        Set Me.pcBlack(ICDraw).Picture = Me.pcCross
                    End If
                Else
                    If BBlack1LastFocus(ICounter) Then
                        Set Me.pcBlack(ICDraw).Picture = Me.pcTickLast
                    Else
                        Set Me.pcBlack(ICDraw).Picture = Me.pcTick
                    End If
                End If
            End If
            
            ICDraw = ICDraw + 1
        Next ICounter
    End If
End Sub

Private Sub SetBoardPattern(ByVal IRoll As Integer, Optional ByVal BAlert As Boolean = True)
    SetNumberRollBoard IRoll
    
    ReDim Preserve IRollB1Pattern(UBound(INumberRollBoard)) As Integer
    ReDim Preserve IRollB2Pattern(UBound(INumberRollBoard)) As Integer
    ReDim Preserve IRollB3Pattern(UBound(INumberRollBoard)) As Integer
    ReDim Preserve IRollB4Pattern(UBound(INumberRollBoard)) As Integer
    
    IRollB1Pattern(UBound(IRollB1Pattern)) = IBlank
    IRollB2Pattern(UBound(IRollB2Pattern)) = IBlank
    IRollB3Pattern(UBound(IRollB3Pattern)) = IBlank
    IRollB4Pattern(UBound(IRollB4Pattern)) = IBlank
    
     If LColorPattern(IRoll) = LRed Then
        SetBlack1RollBoard ICross
        SetBlack1Pattern ICross
        
        SetBlack1Box LOSS
    ElseIf LColorPattern(IRoll) = LBlack Then
        SetBlack1RollBoard ITick
        SetBlack1Pattern ITick
        
        SetBlack1Box WIN
    ElseIf LColorPattern(IRoll) = LGreen Then
        SetBlack1RollBoard INone
        SetBlack1Pattern INone
        
        SetBlack1Box WIN
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
        
        SPattern = Split(mdPattern.CheckPattern(BAlert), vbCrLf)
        
        If BAlert Then
            For ICounter = 0 To Me.lbBet.Count - 1
                Me.lbBet(ICounter).Caption = SPattern(ICounter)
            Next ICounter
        End If
    End If
End Sub

Private Sub SetClear()
    Dim ICounter As Integer
    
    For ICounter = 0 To IRowMax - 1
        Me.lbRoll(ICounter).Caption = ""
        
        Set Me.pcBlack(ICounter).Picture = LoadPicture
        Set Me.pcBlack2(ICounter).Picture = LoadPicture
        Set Me.pcBlack3(ICounter).Picture = LoadPicture
        Set Me.pcBlack4(ICounter).Picture = LoadPicture
    Next ICounter
    
    InitRollBoard
    InitRollPattern
    
    SetBetText
    SetRollFocus
    
    ICounterRollDec = IMax
    
    Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard), "#,##0")
    Me.lbCount(1).Caption = "$ " & Format(ICounterRollDec, "#,##0")
End Sub

Private Sub SetDelete()
    Dim IBlack1RollTemp As Integer
    Dim IBlack2RollTemp As Integer
    Dim IBlack3RollTemp As Integer
    Dim IBlack4RollTemp As Integer
    
    If UBound(INumberRollBoard) <= 0 Then
        INumberRollBoard(0) = IBlank
        IBlack1RollBoard(0) = IBlank
        IBlack2RollBoard(0) = IBlank
        IBlack3RollBoard(0) = IBlank
        IBlack4RollBoard(0) = IBlank
        IBlack1RollPattern(0) = IBlank
        IBlack2RollPattern(0) = IBlank
        IBlack3RollPattern(0) = IBlank
        IBlack4RollPattern(0) = IBlank
        
        Me.lbRoll(0).Caption = ""
        
        Set Me.pcBlack(0).Picture = LoadPicture
        Set Me.pcBlack2(0).Picture = LoadPicture
        Set Me.pcBlack3(0).Picture = LoadPicture
        Set Me.pcBlack4(0).Picture = LoadPicture
        
        SetBetText
        
        BBlack1LastFocus(0) = False
        BBlack2LastFocus(0) = False
        BBlack3LastFocus(0) = False
        BBlack4LastFocus(0) = False
    Else
        IBlack1RollTemp = IBlack1RollBoard(UBound(IBlack1RollBoard))
        IBlack2RollTemp = IBlack2RollBoard(UBound(IBlack2RollBoard))
        IBlack3RollTemp = IBlack3RollBoard(UBound(IBlack3RollBoard))
        IBlack4RollTemp = IBlack4RollBoard(UBound(IBlack3RollBoard))
        
        BBlack1LastFocus(UBound(INumberRollBoard)) = False
        BBlack2LastFocus(UBound(INumberRollBoard)) = False
        BBlack3LastFocus(UBound(INumberRollBoard)) = False
        BBlack4LastFocus(UBound(INumberRollBoard)) = False
        
        ReDim Preserve INumberRollBoard(UBound(INumberRollBoard) - 1) As Integer
        ReDim Preserve IBlack1RollBoard(UBound(IBlack1RollBoard) - 1) As Integer
        ReDim Preserve IBlack1RollPattern(UBound(IBlack1RollPattern) - 1) As Integer
        ReDim Preserve IRollB1Pattern(UBound(INumberRollBoard)) As Integer
        ReDim Preserve IRollB2Pattern(UBound(INumberRollBoard)) As Integer
        ReDim Preserve IRollB3Pattern(UBound(INumberRollBoard)) As Integer
        ReDim Preserve IRollB4Pattern(UBound(INumberRollBoard)) As Integer
        
        If UBound(INumberRollBoard) >= (IRowMax) Then
            SetRollBoard
        Else
            Dim IRCounter As Integer
            Dim ICounter As Integer
            
            IRCounter = (IRowMax - 1) - (IRowMax - UBound(INumberRollBoard) - 1)
            
            For ICounter = 0 To IRowMax - 1
                If ICounter > IRCounter Then
                    Me.lbRoll(ICounter).Caption = ""
                    
                    Set Me.pcBlack(ICounter).Picture = LoadPicture
                    Set Me.pcBlack2(ICounter).Picture = LoadPicture
                    Set Me.pcBlack3(ICounter).Picture = LoadPicture
                    Set Me.pcBlack4(ICounter).Picture = LoadPicture
                Else
                    Me.lbRoll(ICounter).ForeColor = LColorPattern(INumberRollBoard(ICounter))
                    Me.lbRoll(ICounter).Caption = CStr(INumberRollBoard(ICounter))
            
                    If INumberRollBoard(ICounter) = INone Then
                        Set Me.pcBlack(ICounter).Picture = Me.pcNone
                        Set Me.pcBlack2(ICounter).Picture = Me.pcNone
                        Set Me.pcBlack3(ICounter).Picture = Me.pcNone
                        Set Me.pcBlack4(ICounter).Picture = Me.pcNone
                    Else
                        If LColorPattern(INumberRollBoard(ICounter)) = LRed Then
                            Set Me.pcBlack(ICounter).Picture = Me.pcCross
                        Else
                            Set Me.pcBlack(ICounter).Picture = Me.pcTick
                        End If
                    End If
                End If
            Next ICounter
        End If
        
        SetBetText False
        
        CheckSubtWinLoss IBlack1RollTemp, IBlack2RollTemp, IBlack3RollTemp, IBlack4RollTemp
        
        If IBlack1RollTemp = ITick Then
        ElseIf IBlack1RollTemp = INone Then
        Else
        End If
    End If
    
    If INumberRollBoard(UBound(INumberRollBoard)) = IBlank Then
        Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard), "#,##0")
        
        ICounterRollDec = IMax
        
        SetRefreshWinLoss
    Else
        Me.lbCount(0).Caption = " " & Format(UBound(INumberRollBoard) + 1, "#,##0")
        
        If ICounterRollDec < IMax Then ICounterRollDec = ICounterRollDec + 1
    End If
    
    Me.lbCount(1).Caption = "$ " & Format(ICounterRollDec, "#,##0")
End Sub
