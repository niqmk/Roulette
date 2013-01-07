VERSION 5.00
Begin VB.Form fBack 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CloseForm()
    Unload Me
End Sub

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        fMessage.CloseForm
        
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fBack = Nothing
End Sub

Private Sub SetInitial()
    Me.Left = 0
    Me.Top = 0
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Me.AutoRedraw = True
    Me.Print "Press Esc to exit"
End Sub
