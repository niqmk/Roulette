VERSION 5.00
Begin VB.Form fBox 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "SCORE BOARD"
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   1965
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrLoss 
      Left            =   2280
      Top             =   1080
   End
   Begin VB.Timer tmrWin 
      Left            =   720
      Top             =   1080
   End
   Begin VB.PictureBox pcClose 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3120
      Picture         =   "fBox.frx":0000
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   6
      Left            =   0
      TabIndex        =   15
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   5
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1455
      Index           =   4
      Left            =   1800
      TabIndex        =   13
      Top             =   480
      Width           =   30
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   3
      Left            =   3480
      TabIndex        =   12
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   3520
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   1935
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lnBorder 
      BackColor       =   &H00000000&
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   3495
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
      Left            =   2760
      TabIndex        =   8
      Top             =   600
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
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "SCORE BOARD"
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
      Left            =   0
      TabIndex        =   6
      Top             =   120
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
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
      Left            =   120
      TabIndex        =   2
      Top             =   1200
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
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   600
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
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   810
   End
End
Attribute VB_Name = "fBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IBlink As Integer
Private IWinMax As Integer
Private ILossMax As Integer
Private SnX As Single
Private SnY As Single
Private BMove As Boolean

Public Sub CloseForm()
    Unload Me
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

Private Sub lbWin_Change()
    Me.lbWin.Visible = True
    
    If Not (Trim(Me.lbWin.Caption) = "") Then
        Dim IWin As Integer
        
        IWin = CInt(Me.lbWin.Caption)
        
        If IWin > IWinMax Then
            IWinMax = IWin
            
            Me.lbWinMax.Caption = CStr(IWinMax)
        End If
        
        If (IWin = 37) Then
            mdAPI.Beep 1000, 200
            mdAPI.Beep 5400, 400
            mdAPI.Beep 3000, 600
            
            IBlink = 0
            
            Me.tmrWin.Enabled = True
            
            fMessage.SetShape _
                " START BETTING" & _
                vbCrLf & _
                vbCrLf & _
                vbCrLf & _
                vbCrLf & _
                vbCrLf & _
                vbCrLf & _
                "PLAY RS QUICKY", vbBlack, 100, 12200, 2805
            fMessage.SetTimer 30, 500, 1000, 300
            fMessage.Show vbModeless, Me
        Else
            Me.tmrWin.Enabled = False
        End If
    End If
End Sub

Private Sub lbLoss_Change()
    Me.lbLoss.Visible = True
    
    If Not (Trim(Me.lbLoss.Caption) = "") Then
        Dim ILoss As Integer
        
        ILoss = CInt(Me.lbLoss.Caption)
        
        If ILoss > ILossMax Then
            ILossMax = ILoss
            
            Me.lbLossMax.Caption = CStr(ILossMax)
        End If
        
        If (ILoss = 37) Then
            mdAPI.Beep 1000, 200
            mdAPI.Beep 5400, 400
            mdAPI.Beep 3500, 600
            
            IBlink = 0
            
            Me.tmrLoss.Enabled = True
            
            fMessage.Show vbModeless, Me
        Else
            Me.tmrLoss.Enabled = False
        End If
    End If
End Sub

Private Sub pcClose_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fMain.BBox = False
    
    Set fBox = Nothing
End Sub

Private Sub tmrWin_Timer()
    Me.lbWin.Visible = Not Me.lbWin.Visible
    
    IBlink = IBlink + 1
    
    If IBlink > 20 Then
        Me.lbWin.Visible = True
        
        Me.tmrWin.Enabled = False
    End If
End Sub

Private Sub tmrLoss_Timer()
    Me.lbLoss.Visible = Not Me.lbLoss.Visible
    
    IBlink = IBlink + 1
    
    If IBlink > 20 Then
        Me.lbLoss.Visible = True
        
        Me.tmrLoss.Enabled = False
    End If
End Sub

Private Sub SetInitial()
    mdGeneral.CenterWindows Me, True
    
    Me.lbWin.Caption = "0"
    Me.lbWinMax.Caption = "0"
    Me.lbLoss.Caption = "0"
    Me.lbLossMax.Caption = "0"
    
    IWinMax = 0
    ILossMax = 0
    
    With Me.tmrWin
        .Interval = 300
        .Enabled = False
    End With
    
    With Me.tmrLoss
        .Interval = 300
        .Enabled = False
    End With
End Sub
