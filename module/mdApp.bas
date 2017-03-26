Attribute VB_Name = "mdApp"
Option Explicit

Public Enum WinLossBox
    WIN
    LOSS
End Enum

Public Const LRed As Long = &HFF&
Public Const LLRed As Long = &H8080FF
Public Const LBlack As Long = &H0&
Public Const LGreen As Long = &HC000&
Public Const LLGreen As Long = &H80FF80
Public Const LBlue As Long = &HFF0000
Public Const LLBlue As Long = &HFFFF80

Public Const SBlackText As String = "BLACK"
Public Const SRedText As String = "RED"

Public Const IBlank As Integer = -1
Public Const IRowMax As Integer = 20
Public Const INumberMax As Integer = 36
Public Const IMax As Integer = 450

Public Const INone As Integer = 0
Public Const ITick As Integer = 1
Public Const ICross As Integer = 2

Public Const IBlackType As Integer = 1

Public LColorPattern() As Long
Public INumberRollBoard() As Integer
Public IBlack1RollBoard() As Integer
Public IBlack2RollBoard() As Integer
Public IBlack3RollBoard() As Integer
Public IBlack4RollBoard() As Integer
Public IBlack5RollBoard() As Integer
Public IBlack6RollBoard() As Integer
Public IBlack1RollPattern() As Integer
Public IBlack2RollPattern() As Integer
Public IBlack3RollPattern() As Integer
Public IBlack4RollPattern() As Integer
Public IBlack5RollPattern() As Integer
Public IBlack6RollPattern() As Integer
Public IRollB1Pattern() As Integer
Public IRollB2Pattern() As Integer
Public IRollB3Pattern() As Integer
Public IRollB4Pattern() As Integer
Public ICounterRollDec As Integer

Public BBlack1LastFocus(65535) As Boolean
Public BBlack2LastFocus(65335) As Boolean
Public BBlack3LastFocus(65335) As Boolean
Public BBlack4LastFocus(65335) As Boolean
Public BBlack5LastFocus(65335) As Boolean
Public BBlack6LastFocus(65335) As Boolean

Public IBlack1Box(11) As Integer
Public IBlack2Box(11) As Integer
Public IBlack3Box(11) As Integer
Public IBlack4Box(11) As Integer
Public IBlack5Box(11) As Integer
Public IBlack6Box(11) As Integer

Public IBlack1Score As Integer
Public IBlack2Score As Integer
Public IBlack3Score As Integer
Public IBlack4Score As Integer
Public IBlack5Score As Integer
Public IBlack6Score As Integer

Public IBlack1Focus As Integer

Public BBlack1 As Boolean

Private BBlack1Pattern As Boolean
Private BBlack2Pattern As Boolean
Private BBlack3Pattern As Boolean
Private BBlack4Pattern As Boolean
Private BBlack5Pattern As Boolean
Private BBlack6Pattern As Boolean

Private IWin As Integer
Private ILoss As Integer

Public Sub Init()
    InitColorPattern
    InitRollBoard
    InitRollPattern
    
    ICounterRollDec = IMax
    
    IBlack1Score = IBlank
    IBlack2Score = IBlank
    IBlack3Score = IBlank
    IBlack4Score = IBlank
    IBlack5Score = IBlank
    IBlack6Score = IBlank
    
    IBlack1Focus = 0
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
    ReDim IBlack1RollBoard(0) As Integer
    ReDim IBlack2RollBoard(0) As Integer
    ReDim IBlack3RollBoard(0) As Integer
    ReDim IBlack4RollBoard(0) As Integer
    
    INumberRollBoard(0) = IBlank
    IBlack1RollBoard(0) = IBlank
    IBlack2RollBoard(0) = IBlank
    IBlack3RollBoard(0) = IBlank
    IBlack4RollBoard(0) = IBlank
End Sub

Public Sub InitRollPattern()
    ReDim IBlack1RollPattern(0) As Integer
    ReDim IBlack2RollPattern(0) As Integer
    ReDim IBlack3RollPattern(0) As Integer
    ReDim IBlack4RollPattern(0) As Integer
    
    IBlack1RollPattern(0) = IBlank
    IBlack2RollPattern(0) = IBlank
    IBlack3RollPattern(0) = IBlank
    IBlack4RollPattern(0) = IBlank
    
    IWin = 0
    ILoss = 0
    
    BBlack1 = False
End Sub

Public Sub SetBlack1RollBoard(ByVal IValue As Integer)
    If Not (IBlack1RollBoard(UBound(IBlack1RollBoard)) = IBlank) Then
        ReDim Preserve IBlack1RollBoard(UBound(IBlack1RollBoard) + 1) As Integer
    End If
    
    IBlack1RollBoard(UBound(IBlack1RollBoard)) = IValue
End Sub

Public Sub SetBlack1Pattern(ByVal IValue As Integer)
    If UBound(IBlack1RollPattern) = 0 And (IBlack1RollPattern(UBound(IBlack1RollPattern)) = IBlank) Then
        IBlack1RollPattern(UBound(IBlack1RollPattern)) = IValue
    Else
        ReDim Preserve IBlack1RollPattern(UBound(IBlack1RollPattern) + 1) As Integer
        
        IBlack1RollPattern(UBound(IBlack1RollPattern)) = IValue
    End If
End Sub

Public Sub SetBlack1Box(ByVal IType As WinLossBox)
    If IType = WIN Then
        If IBlack1Box(1) > 0 Then
            IBlack1Box(1) = IBlack1Box(1) - 1
        Else
            IBlack1Box(0) = IBlack1Box(0) + 1
        End If
        
        If IBlack1Box(0) > IBlack1Box(2) Then IBlack1Box(2) = IBlack1Box(0)
    ElseIf IType = LOSS Then
        If IBlack1Box(0) > 0 Then
            IBlack1Box(0) = IBlack1Box(0) - 1
        Else
            IBlack1Box(1) = IBlack1Box(1) + 1
        End If
        
        If IBlack1Box(1) > IBlack1Box(3) Then IBlack1Box(3) = IBlack1Box(1)
    End If
    
    If (IBlack1Box(0) = 0) And (IBlack1Box(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IBlack1Box)
            IBlack1Box(ICounter) = 0
        Next ICounter
    Else
        IBlack1Box(4) = IBlack1Box(4) + 1
    End If
End Sub

Public Function CheckFocus(ByVal IType As Integer) As Integer
    If IType = IBlackType Then
        CheckFocus = IBlack1Focus
    End If
End Function

Public Function CheckBlack() As Boolean
    CheckBlack = BBlack1
End Function

Public Sub SetNumberFocus(ByVal IType As Integer)
    If IType = IBlackType Then
        IBlack1Focus = IBlack1Focus + 1
    End If
End Sub

Private Sub SetPatternResult(ByVal SValue As String, ByVal IType As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False, Optional ByVal BSkip As Boolean = False)
    If Not (Trim(SValue) = "") Or BSkip Then
        If BFocus Then SetNumberFocus IType
        If BWinLoss Then SetWinLoss True
    End If
End Sub

Public Sub SetWinLoss(Optional ByVal BWin As Boolean = False)
    If BWin Then
        If ILoss > 0 Then
            ILoss = ILoss - 1
        Else
            IWin = IWin + 1
        End If
    Else
        If IWin > 0 Then
            IWin = IWin - 1
        Else
            ILoss = ILoss + 1
        End If
    End If
End Sub

Public Sub SetWin(Optional ByVal IValue As Integer = 0)
    IWin = IValue
End Sub

Public Sub SetLoss(Optional ByVal IValue As Integer = 0)
    ILoss = IValue
End Sub

Public Sub CheckSubtWinLoss(ByVal IBlack1RollTemp As Integer, ByVal IBlack2RollTemp As Integer, ByVal IBlack3RollTemp As Integer, ByVal IBlack4RollTemp As Integer)
    If CheckBlack Then
        If IBlack1RollTemp = INone Then
            If ILoss > 0 Then
                ILoss = ILoss - 1
            Else
                IWin = IWin + 1
            End If
        Else
            If BBlack1Pattern Then
                If IBlack1RollTemp = ITick Then
                    If IWin > 0 Then
                        IWin = IWin - 1
                    Else
                        ILoss = ILoss + 1
                    End If
                Else
                    If ILoss > 0 Then
                        ILoss = ILoss - 1
                    Else
                        IWin = IWin + 1
                    End If
                End If
            Else
                If IBlack1RollTemp = ICross Then
                    If IWin > 0 Then
                        IWin = IWin - 1
                    Else
                        ILoss = ILoss + 1
                    End If
                Else
                    If ILoss > 0 Then
                        ILoss = ILoss - 1
                    Else
                        IWin = IWin + 1
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub SetRefreshWinLoss()
    SetWin
    SetLoss
    
    If UBound(INumberRollBoard) > 5 Then Exit Sub
    
    Dim IPattern(5) As Integer
    Dim ICounter As Integer
    
    For ICounter = UBound(INumberRollBoard) To LBound(INumberRollBoard) Step -1
        If ICounter > 5 Then
            IPattern(5) = IBlack1RollPattern(ICounter)
            IPattern(4) = IBlack1RollPattern(ICounter - 1)
            IPattern(3) = IBlack1RollPattern(ICounter - 2)
            IPattern(2) = IBlack1RollPattern(ICounter - 3)
            IPattern(1) = IBlack1RollPattern(ICounter - 4)
            IPattern(0) = IBlack1RollPattern(ICounter - 5)
        
            mdPattern.CheckProbPattern IPattern, False, False, True
        Else
            Exit For
        End If
    Next ICounter
End Sub
