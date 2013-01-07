VERSION 5.00
Begin VB.Form fTimer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ScaleHeight     =   990
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcFrame 
      Height          =   975
      Index           =   0
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox pcFrame 
         BackColor       =   &H00000000&
         Height          =   735
         Index           =   1
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   3075
         TabIndex        =   1
         Top             =   120
         Width           =   3135
         Begin VB.Label lbHour 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   465
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lbMinute 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   465
            Left            =   1320
            TabIndex        =   5
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lbSecond 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   465
            Left            =   2520
            TabIndex        =   4
            Top             =   120
            Width           =   480
         End
         Begin VB.Label lbSeparator 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   465
            Index           =   0
            Left            =   840
            TabIndex        =   3
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lbSeparator 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   465
            Index           =   1
            Left            =   2040
            TabIndex        =   2
            Top             =   120
            Width           =   240
         End
      End
   End
   Begin VB.Timer trCount 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "fTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fTimer = Nothing
End Sub

Private Sub trCount_Timer()
    Dim SValue As String
    
    SValue = DateDiff("s", Now, mdSecurity.SGuest)
    
    If SValue < (60 * 5) Then
        Me.lbHour.ForeColor = vbRed
        Me.lbSeparator(0).ForeColor = vbRed
        Me.lbMinute.ForeColor = vbRed
        Me.lbSeparator(1).ForeColor = vbRed
        Me.lbSecond.ForeColor = vbRed
    Else
        Me.lbHour.ForeColor = vbGreen
        Me.lbSeparator(0).ForeColor = vbGreen
        Me.lbMinute.ForeColor = vbGreen
        Me.lbSeparator(1).ForeColor = vbGreen
        Me.lbSecond.ForeColor = vbGreen
    End If
    
    If SValue <= 0 Then
        Me.trCount.Enabled = False
        
        fMain.CloseForm
        
        Exit Sub
    End If
    
    Me.lbHour.Caption = Format(SValue \ 3600, "00")
    SValue = SValue Mod 3600
    Me.lbMinute.Caption = Format(SValue \ 60, "00")
    SValue = SValue Mod 60
    Me.lbSecond.Caption = Format(SValue, "00")
End Sub

Private Sub SetInitial()
    mdGeneral.CenterWindows Me, True, True
    
    mdAPI.SetWindowPos hWnd, mdAPI.HWND_TOPMOST, 0, 0, 0, 0, mdAPI.SWP_NOMOVE + mdAPI.SWP_NOSIZE
    
    Dim SValue As String
    
    SValue = DateDiff("s", Now, mdSecurity.SGuest)
    
    If SValue < (60 * 5) Then
        Me.lbHour.ForeColor = vbRed
        Me.lbSeparator(0).ForeColor = vbRed
        Me.lbMinute.ForeColor = vbRed
        Me.lbSeparator(1).ForeColor = vbRed
        Me.lbSecond.ForeColor = vbRed
    Else
        Me.lbHour.ForeColor = vbGreen
        Me.lbSeparator(0).ForeColor = vbGreen
        Me.lbMinute.ForeColor = vbGreen
        Me.lbSeparator(1).ForeColor = vbGreen
        Me.lbSecond.ForeColor = vbGreen
    End If
    
    Me.lbHour.Caption = Format(SValue \ 3600, "00")
    SValue = SValue Mod 3600
    Me.lbMinute.Caption = Format(SValue \ 60, "00")
    SValue = SValue Mod 60
    Me.lbSecond.Caption = Format(SValue, "00")
    
    Me.trCount.Interval = 1000
    Me.trCount.Enabled = True
End Sub

