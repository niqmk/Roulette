VERSION 5.00
Begin VB.Form fCPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "fCPassword.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMain 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lbMain 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   810
      End
   End
End
Attribute VB_Name = "fCPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Me.txPassword(0).SetFocus
End Sub

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub txPassword_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Index = 2) And (KeyCode = vbKeyReturn) Then
        CheckPassword
    ElseIf KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fCPassword = Nothing
End Sub

Private Sub SetInitial()
    mdMain.BLogin = False
End Sub

Private Sub CheckPassword()
    If mdSecurity.EncryptText(Me.txPassword(0).Text, mdSecurity.SKey) = mdSecurity.SSecured Then
        If Not (Me.txPassword(1).Text = Me.txPassword(2).Text) Then
            MsgBox "Password is not same", vbCritical, mdApp.STitle
            
            Me.txPassword(1).SetFocus
            
            Exit Sub
        End If
        
        mdSecurity.ChangePassword Me.txPassword(1).Text
    
        Unload Me
    Else
        MsgBox "Password is not correct", vbCritical, mdApp.STitle
        
        Me.txPassword(0).Text = ""
        Me.txPassword(0).SetFocus
    End If
End Sub
