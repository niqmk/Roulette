VERSION 5.00
Begin VB.Form fMessage 
   BorderStyle     =   0  'None
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   FillStyle       =   0  'Solid
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMessage 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "fMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lDFHeight As Long
Private lDFColor As Long

Private ICBlink As Integer
Private IDBlink As Integer
Private IDInterval As Integer
Private IDBeep As Integer
Private IDBeepInterval As Integer

Private sDWidth As Single
Private sDHeight As Single
Private sDX As Single
Private sDY As Single

Private SDText As String

Private BDBlink As Boolean

Public Sub CloseForm()
    Unload Me
End Sub

Public Sub SetShape(ByVal SText As String, Optional ByVal lFColor As Long = vbBlack, Optional ByVal lFHeight As Long = 100, Optional ByVal sWidth As Single = 0, Optional ByVal sHeight As Single = 0, Optional ByVal sX As Single = 10, Optional ByVal sY As Single = 10)
    SDText = SText
    lDFColor = lFColor
    lDFHeight = lFHeight
    sDWidth = sWidth
    sDHeight = sHeight
    sDX = sX
    sDY = sY
End Sub

Public Sub SetTimer(ByVal IBlink As Integer, ByVal IInterval As Integer, ByVal IBeep As Integer, ByVal IBeepInterval As Integer, Optional ByVal BBlink As Boolean = True)
    IDBlink = IBlink
    IDInterval = IInterval
    IDBeep = IBeep
    IDBeepInterval = IBeepInterval
    BDBlink = BBlink
End Sub

Private Sub Form_Load()
    fBack.Show
    fMain.Hide
    fMain.HideAllBox

    Me.Width = sDWidth
    Me.Height = sDHeight
    
    mdGeneral.CenterWindows Me, False

    mdAPI.SetWindowPos hWnd, mdAPI.HWND_TOPMOST, 0, 0, 0, 0, mdAPI.SWP_NOMOVE + mdAPI.SWP_NOSIZE

    mdAPI.ShapeForm Me, SDText, lDFColor, lDFHeight, sDX, sDY
    
    With Me.tmrMessage
        .Interval = IDInterval
        .Enabled = True
    End With
    
    ICBlink = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        fBack.CloseForm
        
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not fBack Is Nothing Then fBack.CloseForm
    
    fMain.Show
    If fMain.BBox Then fBox.Show
    If fMain.BTable1 Then fRoll1.Show
    If fMain.BTable2 Then fRoll2.Show
    If fMain.BHistory Then fHistory.Show
    If fMain.BBlack Then fBlack.Show
    If fMain.BRed Then fRed.Show
    If fMain.BEven Then fEven.Show
    If fMain.BOdd Then fOdd.Show
    If fMain.BHigh Then fHigh.Show
    If fMain.BLow Then fLow.Show
    
    Set fMessage = Nothing
End Sub

Private Sub tmrMessage_Timer()
    If BDBlink Then Me.Visible = Not Me.Visible
    
    ICBlink = ICBlink + 1
    
    mdAPI.Beep IDBeep, IDBeepInterval
    
    If ICBlink > IDBlink Then
        ICBlink = 0
    
        If BDBlink Then Me.Visible = True
        
        Me.tmrMessage.Enabled = False
        
        Unload Me
    End If
End Sub
