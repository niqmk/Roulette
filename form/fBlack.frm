VERSION 5.00
Begin VB.Form fBlack 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "BLACK"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   Icon            =   "fBlack.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   24
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txPassword 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      ScrollBars      =   1  'Horizontal
      TabIndex        =   23
      Top             =   120
      Width           =   2895
   End
   Begin VB.PictureBox pcMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      Picture         =   "fBlack.frx":030A
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lbLossMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   41
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbWinMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   40
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lbWinMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   39
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbLossMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   38
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   13
      Left            =   480
      TabIndex        =   37
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   12
      Left            =   480
      TabIndex        =   36
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   11
      Left            =   480
      TabIndex        =   35
      Top             =   5280
      Width           =   3525
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   10
      Left            =   3960
      TabIndex        =   34
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1455
      Index           =   9
      Left            =   2280
      TabIndex        =   33
      Top             =   3840
      Width           =   30
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   8
      Left            =   480
      TabIndex        =   32
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   7
      Left            =   480
      TabIndex        =   31
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   135
      Index           =   7
      Left            =   480
      TabIndex        =   30
      Top             =   3360
      Width           =   3495
   End
   Begin VB.Label lbBox 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "BET"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   29
      Top             =   3480
      Width           =   3495
   End
   Begin VB.Label lbLoss 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   1
      Left            =   2400
      TabIndex        =   28
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lbWin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   600
      TabIndex        =   27
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOSS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   2400
      TabIndex        =   26
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WINS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1320
      TabIndex        =   25
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   6
      Left            =   480
      TabIndex        =   22
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   5
      Left            =   480
      TabIndex        =   21
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1455
      Index           =   4
      Left            =   2280
      TabIndex        =   20
      Top             =   1200
      Width           =   30
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   3
      Left            =   3960
      TabIndex        =   19
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   2640
      Width           =   3520
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   720
      Width           =   45
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label lbCount 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lbCount 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WINS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label lbCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOSS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2400
      TabIndex        =   12
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label lbWin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   600
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lbLoss 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   15
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   5415
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   15
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   4455
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   5415
      Index           =   3
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   135
      Index           =   4
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "BLACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lbBox 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "BLACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   135
      Index           =   6
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   3495
   End
End
Attribute VB_Name = "fBlack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SnX As Single
Private SnY As Single
Private BMove As Boolean

Public Sub CloseForm()
    Unload Me
End Sub

Public Sub SetText()
    Me.lbWin(0).Caption = CStr(mdApp.IBlackBox(0))
    Me.lbLoss(0).Caption = CStr(mdApp.IBlackBox(1))
    Me.lbWinMax(0).Caption = CStr(mdApp.IBlackBox(2))
    Me.lbLossMax(0).Caption = CStr(mdApp.IBlackBox(3))
    Me.lbWin(1).Caption = CStr(mdApp.IBlackBox(5))
    Me.lbLoss(1).Caption = CStr(mdApp.IBlackBox(6))
    Me.lbWinMax(1).Caption = CStr(mdApp.IBlackBox(7))
    Me.lbLossMax(1).Caption = CStr(mdApp.IBlackBox(8))
    Me.lbCount(0).Caption = " " & Format(mdApp.IBlackBox(4), "#,##0")
    Me.lbCount(1).Caption = " " & Format(mdApp.IBoxBaseDec - mdApp.IBlackBox(4), "#,##0")
End Sub

Private Sub Form_Load()
    SetInitial
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BMove = True
    
    SnX = X
    SnY = Y
    
    Me.MousePointer = MousePointerConstants.vbSizeAll
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

Private Sub Form_Unload(Cancel As Integer)
    fMain.BBlack = False
    
    Set fBlack = Nothing
End Sub

Private Sub pcMin_Click()
    Unload Me
End Sub

Private Sub SetInitial()
    mdGeneral.CenterWindows Me, False
    
    Me.txPassword.Visible = False
    
    Me.SetText
End Sub

Private Sub SetClear()
    If mdSecurity.EncryptText(Me.txPassword.Text, mdSecurity.SKey) = mdSecurity.SSecured Then
        fMain.ClearBox
    End If
    
    Me.txPassword.Visible = False
End Sub
