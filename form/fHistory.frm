VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fHistory 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "History"
   ClientHeight    =   9810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton cdDelete 
      Caption         =   "Delete All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvHistory 
      Height          =   8055
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   14208
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.PictureBox pcMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5040
      Picture         =   "fHistory.frx":0000
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   15
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   15
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   9735
      Index           =   3
      Left            =   5400
      TabIndex        =   5
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   9735
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15
   End
   Begin VB.Label lbTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   " HISTORY"
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
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   5415
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
      Width           =   5415
   End
   Begin VB.Label lbBorder 
      BackColor       =   &H00000000&
      Height          =   15
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   5415
   End
End
Attribute VB_Name = "fHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SnX As Single
Private SnY As Single
Private BMove As Boolean
Private BClear As Boolean
Private BPassword As Boolean

Public Sub CloseForm()
    Unload Me
End Sub

Public Sub AddRoll(ByVal IValue As Integer)
    Dim liHistory As ListItem
    
    Set liHistory = Me.lvHistory.ListItems.Add(, , Me.lvHistory.ListItems.Count + 1)
    liHistory.ListSubItems.Add , , IValue
    liHistory.ListSubItems.Add
    liHistory.ListSubItems.Add
    liHistory.ListSubItems.Add
    lvHistory.ListItems(Me.lvHistory.ListItems.Count).ListSubItems(1).ForeColor = mdApp.LColorPattern(IValue)
End Sub

Public Sub DeleteItem()
    mdApp.DeleteHistory
    
    If Me.lvHistory.ListItems.Count > 0 Then Me.lvHistory.ListItems.Remove Me.lvHistory.ListItems.Count
End Sub

Public Sub FillPattern(ByVal IFocus As Integer, ByVal IType As Integer)
    If IFocus >= 0 Then
        Dim liHistory As ListItem
        Dim LColor As Long
        Dim ICounter As Integer
        
        If (Me.lvHistory.ListItems.Count - IFocus) = 6 Then
            LColor = mdApp.LRed
        ElseIf (Me.lvHistory.ListItems.Count - IFocus) = 7 Then
            LColor = mdApp.LBlack
        ElseIf (Me.lvHistory.ListItems.Count - IFocus) > 7 Then
            LColor = mdApp.LLBlue
        End If
        
        For ICounter = Me.lvHistory.ListItems.Count To IFocus + 1 Step -1
            Set liHistory = Me.lvHistory.ListItems(ICounter)
            
            If IType = mdApp.IBlackType Then
                If mdApp.LColorPattern(CInt(liHistory.ListSubItems(1).Text)) = mdApp.LGreen Then
                    liHistory.ListSubItems(2).Text = Chr(45)
                ElseIf mdApp.LColorPattern(CInt(liHistory.ListSubItems(1).Text)) = mdApp.LRed Then
                    liHistory.ListSubItems(2).Text = Chr(47)
                ElseIf mdApp.LColorPattern(CInt(liHistory.ListSubItems(1).Text)) = mdApp.LBlack Then
                    liHistory.ListSubItems(2).Text = Chr(120)
                End If
                
                liHistory.ListSubItems(2).ForeColor = LColor
                liHistory.ListSubItems(2).Bold = True
            ElseIf IType = mdApp.IEvenType Then
                If mdApp.LColorPattern(CInt(liHistory.ListSubItems(1).Text)) = mdApp.LGreen Then
                    liHistory.ListSubItems(3).Text = Chr(45)
                ElseIf CInt(liHistory.ListSubItems(1).Text) Mod 2 = 0 Then
                    liHistory.ListSubItems(3).Text = Chr(120)
                Else
                    liHistory.ListSubItems(3).Text = Chr(47)
                End If
                
                liHistory.ListSubItems(3).ForeColor = LColor
                liHistory.ListSubItems(3).Bold = True
            ElseIf IType = mdApp.IHighType Then
                If mdApp.LColorPattern(CInt(liHistory.ListSubItems(1).Text)) = mdApp.LGreen Then
                    liHistory.ListSubItems(4).Text = Chr(45)
                ElseIf CInt(liHistory.ListSubItems(1).Text) >= 19 Then
                    liHistory.ListSubItems(4).Text = Chr(120)
                Else
                    liHistory.ListSubItems(4).Text = Chr(47)
                End If
                
                liHistory.ListSubItems(4).ForeColor = LColor
                liHistory.ListSubItems(4).Bold = True
            End If
        Next ICounter
    End If
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

Private Sub Form_Unload(Cancel As Integer)
    fMain.BHistory = False
    
    Set fHistory = Nothing
End Sub

Private Sub txPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SetDelete
    End If
End Sub

Private Sub txPassword_LostFocus()
    Me.txPassword.Visible = False
End Sub

Private Sub pcMin_Click()
    Unload Me
End Sub

Private Sub cdDelete_Click(Index As Integer)
    If Index = 0 Then
        Me.txPassword.Top = Me.cdDelete(1).Top
        
        BClear = True
    Else
        If BPassword Then
            DeleteItem
            
            Exit Sub
        Else
            Me.txPassword.Top = Me.cdDelete(0).Top
            
            BClear = False
        End If
    End If
    
    Me.txPassword.Text = ""
    Me.txPassword.Visible = True
    Me.txPassword.SetFocus
End Sub

Private Sub SetInitial()
    mdGeneral.CenterWindows Me, False
    
    With Me.lvHistory
        .View = lvwReport
        .ColumnHeaders.Add , , "No.", 1000
        .ColumnHeaders.Add , , "Roll", 1100
        .ColumnHeaders.Add , , , 400, ListColumnAlignmentConstants.lvwColumnCenter
        .ColumnHeaders.Add , , , 400, ListColumnAlignmentConstants.lvwColumnCenter
        .ColumnHeaders.Add , , , 400, ListColumnAlignmentConstants.lvwColumnCenter
    End With
    
    Dim liHistory As ListItem
    Dim ICounter As Integer
    
    If mdApp.IRollHTS(UBound(mdApp.IRollHTS)) = IBlank Then
    Else
        Dim LBColor As Long
        Dim LEColor As Long
        Dim LHColor As Long
        Dim IBValue As Integer
        Dim IEValue As Integer
        Dim IHValue As Integer
        Dim ITemp As Integer
        Dim BBPivot As Boolean
        Dim BEPivot As Boolean
        Dim BHPivot As Boolean
        
        LBColor = mdApp.LBlack
        LEColor = mdApp.LBlack
        LHColor = mdApp.LBlack
        
        BBPivot = False
        BEPivot = False
        BHPivot = False
        
        For ICounter = 0 To UBound(mdApp.IRollHTS)
            Set liHistory = Me.lvHistory.ListItems.Add(, , ICounter + 1)
            liHistory.ListSubItems.Add , , mdApp.IRollHTS(ICounter)
            lvHistory.ListItems(ICounter + 1).ListSubItems(1).ForeColor = mdApp.LColorPattern(mdApp.IRollHTS(ICounter))
            
            If mdApp.IRollBPatternHTS(ICounter) = mdApp.IBlank Then
                LBColor = mdApp.LBlack
                
                BBPivot = False
            Else
                If Not BBPivot Then
                    IBValue = 0
                    
                    For ITemp = ICounter To UBound(mdApp.IRollHTS)
                        If mdApp.IRollBPatternHTS(ITemp) = mdApp.IBlank Then
                            Exit For
                        Else
                            IBValue = IBValue + 1
                        End If
                    Next ITemp
                    
                    If IBValue = 6 Then
                        LBColor = mdApp.LRed
                    ElseIf IBValue = 7 Then
                        LBColor = mdApp.LBlack
                    ElseIf IBValue > 7 Then
                        LBColor = mdApp.LLBlue
                    End If
                    
                    BBPivot = True
                End If
            End If
            
            If mdApp.IRollBPatternHTS(ICounter) = mdApp.INone Then
                liHistory.ListSubItems.Add , , Chr(45)
            ElseIf mdApp.IRollBPatternHTS(ICounter) = mdApp.ITick Then
                liHistory.ListSubItems.Add , , Chr(47)
            ElseIf mdApp.IRollBPatternHTS(ICounter) = mdApp.ICross Then
                liHistory.ListSubItems.Add , , Chr(120)
            Else
                liHistory.ListSubItems.Add
            End If
            
            liHistory.ListSubItems(2).ForeColor = LBColor
            liHistory.ListSubItems(2).Bold = True
            
            If mdApp.IRollEPatternHTS(ICounter) = mdApp.IBlank Then
                LEColor = mdApp.LBlack
                
                BEPivot = False
            Else
                If Not BEPivot Then
                    IEValue = 0
                    
                    For ITemp = ICounter To UBound(mdApp.IRollHTS)
                        If mdApp.IRollEPatternHTS(ITemp) = mdApp.IBlank Then
                            Exit For
                        Else
                            IEValue = IEValue + 1
                        End If
                    Next ITemp
                    
                    If IEValue = 6 Then
                        LEColor = mdApp.LRed
                    ElseIf IEValue = 7 Then
                        LEColor = mdApp.LBlack
                    ElseIf IEValue > 7 Then
                        LEColor = mdApp.LLBlue
                    End If
                    
                    BEPivot = True
                End If
            End If
            
            If mdApp.IRollEPatternHTS(ICounter) = mdApp.INone Then
                liHistory.ListSubItems.Add , , Chr(45)
            ElseIf mdApp.IRollEPatternHTS(ICounter) = mdApp.ITick Then
                liHistory.ListSubItems.Add , , Chr(47)
            ElseIf mdApp.IRollEPatternHTS(ICounter) = mdApp.ICross Then
                liHistory.ListSubItems.Add , , Chr(120)
            Else
                liHistory.ListSubItems.Add
            End If
            
            liHistory.ListSubItems(3).ForeColor = LEColor
            liHistory.ListSubItems(3).Bold = True
            
            If mdApp.IRollHPatternHTS(ICounter) = mdApp.IBlank Then
                LHColor = mdApp.LBlack
                
                BHPivot = False
            Else
                If Not BHPivot Then
                    IHValue = 0
                    
                    For ITemp = ICounter To UBound(mdApp.IRollHTS)
                        If mdApp.IRollHPatternHTS(ITemp) = mdApp.IBlank Then
                            Exit For
                        Else
                            IHValue = IHValue + 1
                        End If
                    Next ITemp
                    
                    If IHValue = 6 Then
                        LHColor = mdApp.LRed
                    ElseIf IHValue = 7 Then
                        LHColor = mdApp.LBlack
                    ElseIf IHValue > 7 Then
                        LHColor = mdApp.LLBlue
                    End If
                    
                    BHPivot = True
                End If
            End If
            
            If mdApp.IRollHPatternHTS(ICounter) = mdApp.INone Then
                liHistory.ListSubItems.Add , , Chr(45)
            ElseIf mdApp.IRollHPatternHTS(ICounter) = mdApp.ITick Then
                liHistory.ListSubItems.Add , , Chr(47)
            ElseIf mdApp.IRollHPatternHTS(ICounter) = mdApp.ICross Then
                liHistory.ListSubItems.Add , , Chr(120)
            Else
                liHistory.ListSubItems.Add
            End If
            
            liHistory.ListSubItems(4).ForeColor = LHColor
            liHistory.ListSubItems(4).Bold = True
        Next ICounter
    End If
    
    Me.txPassword.Visible = False
    BClear = True
    BPassword = False
End Sub

Private Sub SetDelete()
    If BClear Then
        If mdSecurity.EncryptText(Me.txPassword.Text, mdSecurity.SKey) = mdSecurity.SSecured Then
            mdApp.ClearHistory
            
            Me.lvHistory.ListItems.Clear
        End If
    Else
        If mdSecurity.EncryptText(Me.txPassword.Text, mdSecurity.SKey) = mdSecurity.SSecured Then
            mdApp.DeleteHistory
            
            Me.lvHistory.ListItems.Remove Me.lvHistory.ListItems.Count
            
            BPassword = True
        End If
    End If
    
    Me.txPassword.Visible = False
End Sub
