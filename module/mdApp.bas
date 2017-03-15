Attribute VB_Name = "mdApp"
Option Explicit

Public Const LRed As Long = &HFF&
Public Const LLRed As Long = &H8080FF
Public Const LBlack As Long = &H0&
Public Const LGreen As Long = &HC000&
Public Const LLGreen As Long = &H80FF80
Public Const LBlue As Long = &HFF0000
Public Const LLBlue As Long = &HFFFF80

Public Const IBlank As Integer = 0
Public Const IRowMax As Integer = 20
Public Const INumberMax As Integer = 36
Public Const IMax As Integer = 450

Public LColorPattern() As Long
Public INumberRollBoard() As Integer
Public ICounterRollDec As Integer

Private IBlack1Score As Integer
Private IBlack2Score As Integer
Private IBlack3Score As Integer
Private IBlack4Score As Integer
Private IWin As Integer
Private ILoss As Integer

Public Sub Init()
    InitColorPattern
    InitRollBoard
    
    ICounterRollDec = IMax
    
    IBlack1Score = IBlank
    IBlack2Score = IBlank
    IBlack3Score = IBlank
    IBlack4Score = IBlank
End Sub

Public Sub SetNumberRollBoard(ByVal IValue As Integer, Optional ByVal BTable As Boolean = True)
    If Not (INumberRollBoard(UBound(INumberRollBoard)) = IBlank) Then
        ReDim Preserve INumberRollBoard(UBound(INumberRollBoard) + 1) As Integer
    End If
    
    INumberRollBoard(UBound(INumberRollBoard)) = IValue
End Sub

Private Sub InitColorPattern()
    ReDim LColorPattern(0 To INumberMax) As Long
    
    LColorPattern(0) = LGreen
    LColorPattern(1) = LRed
    LColorPattern(2) = LBlack
    LColorPattern(3) = LRed
    LColorPattern(4) = LBlack
    LColorPattern(5) = LRed
    LColorPattern(6) = LBlack
    LColorPattern(7) = LRed
    LColorPattern(8) = LBlack
    LColorPattern(9) = LRed
    LColorPattern(10) = LBlack
    LColorPattern(11) = LBlack
    LColorPattern(12) = LRed
    LColorPattern(13) = LBlack
    LColorPattern(14) = LRed
    LColorPattern(15) = LBlack
    LColorPattern(16) = LRed
    LColorPattern(17) = LBlack
    LColorPattern(18) = LRed
    LColorPattern(19) = LRed
    LColorPattern(20) = LBlack
    LColorPattern(21) = LRed
    LColorPattern(22) = LBlack
    LColorPattern(23) = LRed
    LColorPattern(24) = LBlack
    LColorPattern(25) = LRed
    LColorPattern(26) = LBlack
    LColorPattern(27) = LRed
    LColorPattern(28) = LBlack
    LColorPattern(29) = LBlack
    LColorPattern(30) = LRed
    LColorPattern(31) = LBlack
    LColorPattern(32) = LRed
    LColorPattern(33) = LBlack
    LColorPattern(34) = LRed
    LColorPattern(35) = LBlack
    LColorPattern(36) = LRed
End Sub

Public Sub InitRollBoard()
    ReDim INumberRollBoard(0) As Integer
    
    INumberRollBoard(0) = IBlank
End Sub
