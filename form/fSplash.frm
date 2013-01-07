VERSION 5.00
Begin VB.Form fSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMain 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.Timer tmMain 
         Left            =   3840
         Top             =   2160
      End
      Begin VB.PictureBox pcMain 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1920
         Left            =   120
         Picture         =   "fSplash.frx":0000
         ScaleHeight     =   1920
         ScaleWidth      =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For 32-bit Windows"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2160
         TabIndex        =   4
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roulette System Quicky"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   2280
         Width           =   825
      End
   End
   Begin VB.Line lnBorder 
      BorderWidth     =   5
      Index           =   3
      X1              =   4440
      X2              =   4440
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line lnBorder 
      BorderWidth     =   5
      Index           =   2
      X1              =   0
      X2              =   4440
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line lnBorder 
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2880
   End
   Begin VB.Line lnBorder 
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   4440
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fSplash = Nothing
End Sub

Private Sub SetInitial()
    With fSplash.tmMain
        .Interval = 2000
        .Enabled = True
    End With
End Sub

Private Sub tmMain_Timer()
    fSplash.tmMain.Enabled = False
    
    Unload Me
End Sub
