VERSION 5.00
Begin VB.Form fLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Roulette System Quicky"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "fLogin.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frMain 
      Height          =   735
      Left            =   120
      TabIndex        =   0
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
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   2775
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   810
      End
   End
End
Attribute VB_Name = "fLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Me.txPassword.SetFocus
End Sub

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fLogin = Nothing
End Sub

Private Sub txPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then CheckLogin
End Sub

Private Sub SetInitial()
    mdMain.BLogin = False
    
    mdSecurity.Init
End Sub

Private Sub CheckLogin()
    If mdSecurity.EncryptText(Me.txPassword.Text, mdSecurity.SKey) = mdSecurity.SSecured Then
        mdMain.BLogin = True
        mdSecurity.IsGuest False
        
        Unload Me
    Else
        If mdSecurity.EncryptText(Me.txPassword.Text, mdSecurity.SKey) = mdSecurity.SDGuest Then
            If mdSecurity.IsGuest Then
                mdMain.BLogin = True
            Else
                mdMain.BLogin = False
            End If
            
            Unload Me
            
            Exit Sub
        End If
        
        mdMain.BLogin = False
        
        MsgBox "Password is not correct", vbCritical, mdApp.STitle
        
        Me.txPassword.Text = ""
        Me.txPassword.SetFocus
    End If
End Sub
