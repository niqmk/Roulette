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

Public Const IRowMax As Integer = 20
Public Const INumberMax As Integer = 36
Public Const IBoxWin As Integer = 37
Public Const IBoxStopBet As Integer = 10
Public Const INone As Integer = 0
Public Const ITick As Integer = 1
Public Const ICross As Integer = 2
Public Const IBlank As Integer = -1
Public Const IBlackType As Integer = 1
Public Const IEvenType As Integer = 2
Public Const IHighType As Integer = 3
Public Const IHistoryMax As Integer = 5000
Public Const IRollBaseDec As Integer = 10000
Public Const IBoxBaseDec As Integer = 1350

Public Const STitle As String = "Roulette System Quicky"

Public LColorPattern() As Long

Public INumberRollBoard() As Integer
Public IBlackRollBoard() As Integer
Public IEvenRollBoard() As Integer
Public IHighRollBoard() As Integer
Public IBlackRollPattern() As Integer
Public IEvenRollPattern() As Integer
Public IHighRollPattern() As Integer
Public IRollHT1() As Integer
Public IRollHT2() As Integer
Public IRollHTS() As Integer
Public IRollBPatternHTS() As Integer
Public IRollEPatternHTS() As Integer
Public IRollHPatternHTS() As Integer

'0 - Win Count
'1 - Loss Count
'2 - Win Max Count
'3 - Loss Max Count
'4 - Counter
'5 - Win Bet Count
'6 - Loss Bet Count
'7 - Win Max Bet Count
'8 - Loss Max Bet Count
'9 - Flag Bet
'10 - Flag Message Bet
'11 - Flag Stop Message Bet
Public IBlackBox(11) As Integer
Public IRedBox(11) As Integer
Public IEvenBox(11) As Integer
Public IOddBox(11) As Integer
Public IHighBox(11) As Integer
Public ILowBox(11) As Integer

'0 - Black
'1 - Red
'2 - Even
'3 - Odd
'4 - High
'5 - Low
Public IRollCHT1(5) As Integer
Public IRollCHT2(5) As Integer
Public IRollWBHT1(5) As Integer
Public IRollLBHT1(5) As Integer
Public IRollWBHT2(5) As Integer
Public IRollLBHT2(5) As Integer

Public BBlackLastFocus(65535) As Boolean
Public BEvenLastFocus(65335) As Boolean
Public BHighLastFocus(65335) As Boolean

Public ICounterRollDec As Integer

Private Const IText1 As Integer = 1
Private Const IText2 As Integer = 2

Public Const SBlackText As String = "BLACK"
Public Const SRedText As String = "RED"
Public Const SEvenText As String = "EVEN"
Public Const SOddText As String = "ODD"
Public Const SHighText As String = "HIGH"
Public Const SLowText As String = "LOW"

Private IBlackScore As Integer
Private IEvenScore As Integer
Private IHighScore As Integer
Private IBlackFocus As Integer
Private IEvenFocus As Integer
Private IHighFocus As Integer
Private IWin As Integer
Private ILoss As Integer
Private ITable As Integer
Private IBOutcome As Integer
Private IEOutcome As Integer
Private IHOutcome As Integer
Private BBlack As Boolean
Private BEven As Boolean
Private BHigh As Boolean
Private BBlackPattern As Boolean
Private BEvenPattern As Boolean
Private BHighPattern As Boolean

Public Sub Init()
    InitColorPattern
    InitRollBoard
    InitRollPattern
    InitRollHistory
    InitHistory
    InitPatternHistory
    InitBox
    
    IBlackScore = IBlank
    IEvenScore = IBlank
    IHighScore = IBlank
    
    IBlackFocus = 0
    IEvenFocus = 0
    IHighFocus = 0
End Sub

Public Function CheckTableHistory() As Integer
    CheckTableHistory = ITable
End Function

Public Function CheckFocus(ByVal IType As Integer) As Integer
    If IType = IBlackType Then
        CheckFocus = IBlackFocus
    ElseIf IType = IEvenType Then
        CheckFocus = IEvenFocus
    ElseIf IType = IHighType Then
        CheckFocus = IHighFocus
    End If
End Function

Public Function CheckBlack() As Boolean
    CheckBlack = BBlack
End Function

Public Function CheckEven() As Boolean
    CheckEven = BEven
End Function

Public Function CheckHigh() As Boolean
    CheckHigh = BHigh
End Function

Public Function CheckWin() As Integer
    CheckWin = IWin
End Function

Public Function CheckLoss() As Integer
    CheckLoss = ILoss
End Function

Public Function CheckBOutcome() As Integer
    CheckBOutcome = IBOutcome
End Function

Public Function CheckEOutcome() As Integer
    CheckEOutcome = IEOutcome
End Function

Public Function CheckHOutcome() As Integer
    CheckHOutcome = IHOutcome
End Function

Public Sub CheckSubtWinLoss(ByVal IBlackRollTemp As Integer, ByVal IEvenRollTemp As Integer, ByVal IHighRollTemp As Integer)
    If CheckBlack Then
        If IBlackRollTemp = INone Then
            If CheckEven Or CheckHigh Then
            Else
                If ILoss > 0 Then
                    ILoss = ILoss - 1
                Else
                    IWin = IWin + 1
                End If
            End If
        Else
            If BBlackPattern Then
                If IBlackRollTemp = ITick Then
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
                If IBlackRollTemp = ICross Then
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
    
    If CheckEven Then
        If IEvenRollTemp = INone Then
            If CheckBlack Or CheckHigh Then
            Else
                If ILoss > 0 Then
                    ILoss = ILoss - 1
                Else
                    IWin = IWin + 1
                End If
            End If
        Else
            If BEvenPattern Then
                If IEvenRollTemp = ITick Then
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
                If IEvenRollTemp = ICross Then
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
    
    If CheckHigh Then
        If IHighRollTemp = INone Then
            If CheckBlack Or CheckHigh Then
            Else
                If ILoss > 0 Then
                    ILoss = ILoss - 1
                Else
                    IWin = IWin + 1
                End If
            End If
        Else
            If BHighPattern Then
                If IHighRollTemp = ITick Then
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
                If IHighRollTemp = ICross Then
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
            IPattern(5) = IBlackRollPattern(ICounter)
            IPattern(4) = IBlackRollPattern(ICounter - 1)
            IPattern(3) = IBlackRollPattern(ICounter - 2)
            IPattern(2) = IBlackRollPattern(ICounter - 3)
            IPattern(1) = IBlackRollPattern(ICounter - 4)
            IPattern(0) = IBlackRollPattern(ICounter - 5)
        
            CheckProbPattern IPattern, IBlackType, False, False, True
            
            IPattern(5) = IEvenRollPattern(ICounter)
            IPattern(4) = IEvenRollPattern(ICounter - 1)
            IPattern(3) = IEvenRollPattern(ICounter - 2)
            IPattern(2) = IEvenRollPattern(ICounter - 3)
            IPattern(1) = IEvenRollPattern(ICounter - 4)
            IPattern(0) = IEvenRollPattern(ICounter - 5)
        
            CheckProbPattern IPattern, IEvenType, False, False, True
            
            IPattern(5) = IHighRollPattern(ICounter)
            IPattern(4) = IHighRollPattern(ICounter - 1)
            IPattern(3) = IHighRollPattern(ICounter - 2)
            IPattern(2) = IHighRollPattern(ICounter - 3)
            IPattern(1) = IHighRollPattern(ICounter - 4)
            IPattern(0) = IHighRollPattern(ICounter - 5)
        
            CheckProbPattern IPattern, IHighType, False, False, True
        Else
            Exit For
        End If
    Next ICounter
End Sub

Public Sub SetTableHistory(Optional ITableHistory As Integer = 1)
    ITable = ITableHistory
    
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, lRegKey)
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE)
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, "Table", CStr(ITable))
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub SaveRollHistory(ByRef IRoll() As Integer, ByRef IRollC() As Integer)
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO)
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, lRegKey)
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE)
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Dim ICounter As Integer
    
    For ICounter = LBound(IRoll) To UBound(IRoll)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, CStr(ICounter) & "-" & CStr(ITable), CStr(IRoll(ICounter)))
    Next ICounter
    
    Dim SValue As String
    
    SValue = ""
    
    For ICounter = LBound(IRollC) To UBound(IRollC)
        If Not (Trim(SValue) = "") Then SValue = SValue & "|"
        
        SValue = SValue & CStr(IRollC(ICounter))
    Next ICounter
    
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, "CNT - " & CStr(ITable), SValue)

    For ICounter = LBound(INumberRollBoard) To UBound(INumberRollBoard)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROW - " & CStr(ICounter + 1), CStr(INumberRollBoard(ICounter)))
    Next ICounter
    
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROLL DEC", CStr(ICounterRollDec))
    
    Exit Sub
ErrHandler:
End Sub

Public Sub SavePatternHistory()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, lRegKey)
    
    If lRegistry = 0 Then
        Dim ICounter As Integer
        Dim SBPatternHTS As String
        Dim SEPatternHTS As String
        Dim SHPatternHTS As String
        
        SBPatternHTS = ""
        SEPatternHTS = ""
        SHPatternHTS = ""
        
        For ICounter = LBound(IRollBPatternHTS) To UBound(IRollBPatternHTS)
            SBPatternHTS = SBPatternHTS & "|" & CStr(IRollBPatternHTS(ICounter))
        Next ICounter
        
        For ICounter = LBound(IRollEPatternHTS) To UBound(IRollEPatternHTS)
            SEPatternHTS = SEPatternHTS & "|" & CStr(IRollEPatternHTS(ICounter))
        Next ICounter
        
        For ICounter = LBound(IRollHPatternHTS) To UBound(IRollHPatternHTS)
            SHPatternHTS = SHPatternHTS & "|" & CStr(IRollHPatternHTS(ICounter))
        Next ICounter
        
        If Not (Trim(SBPatternHTS) = "") Then SBPatternHTS = Mid(SBPatternHTS, 2)
        If Not (Trim(SEPatternHTS) = "") Then SEPatternHTS = Mid(SEPatternHTS, 2)
        If Not (Trim(SHPatternHTS) = "") Then SHPatternHTS = Mid(SHPatternHTS, 2)
        
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, SBlackText, SBPatternHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, SEvenText, SEPatternHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, SHighText, SHPatternHTS)
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub SaveBoxHistory()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, lRegKey)
    
    If lRegistry = 0 Then
        Dim ICounter As Integer
        Dim SBBoxHTS As String
        Dim SRBoxHTS As String
        Dim SEBoxHTS As String
        Dim SOBoxHTS As String
        Dim SHBoxHTS As String
        Dim SLBoxHTS As String
        
        SBBoxHTS = ""
        SRBoxHTS = ""
        SEBoxHTS = ""
        SOBoxHTS = ""
        SHBoxHTS = ""
        SLBoxHTS = ""
        
        For ICounter = LBound(IBlackBox) To UBound(IBlackBox)
            If Not (Trim(SBBoxHTS) = "") Then SBBoxHTS = SBBoxHTS & "|"
            SBBoxHTS = SBBoxHTS & CStr(IBlackBox(ICounter))
        Next ICounter
        
        For ICounter = LBound(IRedBox) To UBound(IRedBox)
            If Not (Trim(SRBoxHTS) = "") Then SRBoxHTS = SRBoxHTS & "|"
            SRBoxHTS = SRBoxHTS & CStr(IRedBox(ICounter))
        Next ICounter
        
        For ICounter = LBound(IEvenBox) To UBound(IEvenBox)
            If Not (Trim(SEBoxHTS) = "") Then SEBoxHTS = SEBoxHTS & "|"
            SEBoxHTS = SEBoxHTS & CStr(IEvenBox(ICounter))
        Next ICounter
        
        For ICounter = LBound(IOddBox) To UBound(IOddBox)
            If Not (Trim(SOBoxHTS) = "") Then SOBoxHTS = SOBoxHTS & "|"
            SOBoxHTS = SOBoxHTS & CStr(IOddBox(ICounter))
        Next ICounter
        
        For ICounter = LBound(IHighBox) To UBound(IHighBox)
            If Not (Trim(SHBoxHTS) = "") Then SHBoxHTS = SHBoxHTS & "|"
            SHBoxHTS = SHBoxHTS & CStr(IHighBox(ICounter))
        Next ICounter
        
        For ICounter = LBound(ILowBox) To UBound(ILowBox)
            If Not (Trim(SLBoxHTS) = "") Then SLBoxHTS = SLBoxHTS & "|"
            SLBoxHTS = SLBoxHTS & CStr(ILowBox(ICounter))
        Next ICounter
        
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, SBlackText, SBBoxHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, SRedText, SRBoxHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, SEvenText, SEBoxHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, SOddText, SOBoxHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, SHighText, SHBoxHTS)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, SLowText, SLBoxHTS)
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub DeleteRollHistory()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO)
    
    Dim ICPivot As Integer
    
    If (UBound(INumberRollBoard) = 0) And (INumberRollBoard(0) = IBlank) Then
        ICPivot = 1
    Else
        ICPivot = UBound(INumberRollBoard) + 2
    End If
    
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROW - " & CStr(ICPivot), CStr(IBlank))
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROLL DEC", CStr(ICounterRollDec))
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub ClearRollHistory(ByRef IRoll() As Integer, ByRef IRollC() As Integer)
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, lRegKey)
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE)
    
    Dim ICounter As Integer
    
    For ICounter = LBound(IRoll) To UBound(IRoll)
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, CStr(ICounter) & "-" & CStr(ITable), "0")
        
        IRoll(ICounter) = 0
    Next ICounter
    
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, "CNT - " & CStr(ITable), "")
    
    For ICounter = LBound(IRollC) To UBound(IRollC)
        IRollC(ICounter) = 0
    Next ICounter
    
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, "CNT - " & CStr(ITable), "0")
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub ClearBoardHistory()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    
    If lRegistry = 0 Then
        Dim SValue As String
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, "ROLL DEC", lType, SValue, lSize)
        
        SValue = mdAPI.ReplaceRegistry(SValue)
        
        If IsNumeric(SValue) Then
            Dim ICounter As Integer
            
            For ICounter = 1 To (IRollBaseDec - CInt(SValue))
                mdAPI.DeleteSubKeysRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROW - " & CStr(ICounter)
            Next ICounter
        End If
        
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, "ROLL DEC", CStr(IRollBaseDec))
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub ClearHistory()
    On Local Error GoTo ErrHandler
    
    ReDim IRollHTS(0) As Integer
    IRollHTS(0) = IBlank
    
    mdAPI.DeleteKeysRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY
    
    Exit Sub
ErrHandler:
End Sub

Public Sub DeleteHistory()
    On Local Error GoTo ErrHandler
    
    If UBound(IRollHTS) > 0 Then
        Dim lRegistry As Long
        Dim lRegKey As Long
        Dim lType As Long
        Dim lSize As Long
        
        lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, lRegKey)
        If lRegistry = 0 Then mdAPI.DeleteSubKeysRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, CStr(UBound(IRollHTS) + 1)
        
        ReDim Preserve IRollHTS(UBound(IRollHTS) - 1) As Integer
    Else
        IRollHTS(0) = IBlank
        
        mdAPI.DeleteKeysRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY
    End If
    
    Exit Sub
ErrHandler:
End Sub

Public Sub DeleteBoxHistory()
    On Local Error GoTo ErrHandler
    
    Dim ICounter As Integer
    
    For ICounter = LBound(IBlackBox) To UBound(IBlackBox)
        IBlackBox(ICounter) = 0
        IRedBox(ICounter) = 0
        IEvenBox(ICounter) = 0
        IOddBox(ICounter) = 0
        IHighBox(ICounter) = 0
        ILowBox(ICounter) = 0
    Next ICounter
    
    mdAPI.DeleteKeysRegistry mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX
    
    Exit Sub
ErrHandler:
End Sub

Public Sub SetNumberRollBoard(ByVal IValue As Integer, Optional ByVal BTable As Boolean = True)
    If Not (INumberRollBoard(UBound(INumberRollBoard)) = IBlank) Then
        ReDim Preserve INumberRollBoard(UBound(INumberRollBoard) + 1) As Integer
    End If
    
    INumberRollBoard(UBound(INumberRollBoard)) = IValue
    
    If BTable Then
        If ITable = 1 Then
            IRollHT1(IValue) = IRollHT1(IValue) + 1
            
            If mdApp.LColorPattern(IValue) = mdApp.LRed Then
                IRollCHT1(1) = IRollCHT1(1) + 1
            ElseIf LColorPattern(IValue) = mdApp.LBlack Then
                IRollCHT1(0) = IRollCHT1(0) + 1
            ElseIf LColorPattern(IValue) = mdApp.LGreen Then
                IRollCHT1(0) = IRollCHT1(0) + 1
                IRollCHT1(1) = IRollCHT1(1) + 1
            End If
            
            If IValue = 0 Then
                IRollCHT1(2) = IRollCHT1(2) + 1
                IRollCHT1(3) = IRollCHT1(3) + 1
            ElseIf IValue Mod 2 = 0 Then
                IRollCHT1(2) = IRollCHT1(2) + 1
            Else
                IRollCHT1(3) = IRollCHT1(3) + 1
            End If
            
            If IValue = 0 Then
                IRollCHT1(4) = IRollCHT1(4) + 1
                IRollCHT1(5) = IRollCHT1(5) + 1
            ElseIf IValue >= 19 Then
                IRollCHT1(4) = IRollCHT1(4) + 1
            Else
                IRollCHT1(5) = IRollCHT1(5) + 1
            End If
            
            SaveRollHistory IRollHT1, IRollCHT1
        ElseIf ITable = 2 Then
            IRollHT2(IValue) = IRollHT2(IValue) + 1
            
            If mdApp.LColorPattern(IValue) = mdApp.LRed Then
                IRollCHT2(1) = IRollCHT2(1) + 1
            ElseIf LColorPattern(IValue) = mdApp.LBlack Then
                IRollCHT2(0) = IRollCHT2(0) + 1
            ElseIf LColorPattern(IValue) = mdApp.LGreen Then
                IRollCHT2(0) = IRollCHT2(0) + 1
                IRollCHT2(1) = IRollCHT2(1) + 1
            End If
            
            If IValue = 0 Then
                IRollCHT2(2) = IRollCHT2(2) + 1
                IRollCHT2(3) = IRollCHT2(3) + 1
            ElseIf IValue Mod 2 = 0 Then
                IRollCHT2(2) = IRollCHT2(2) + 1
            Else
                IRollCHT2(3) = IRollCHT2(3) + 1
            End If
            
            If IValue = 0 Then
                IRollCHT2(4) = IRollCHT2(4) + 1
                IRollCHT2(5) = IRollCHT2(5) + 1
            ElseIf IValue >= 19 Then
                IRollCHT2(4) = IRollCHT2(4) + 1
            Else
                IRollCHT2(5) = IRollCHT2(5) + 1
            End If
            
            SaveRollHistory IRollHT2, IRollCHT2
        End If
    End If
End Sub

Public Sub SetBlackRollBoard(ByVal IValue As Integer)
    If Not (IBlackRollBoard(UBound(IBlackRollBoard)) = IBlank) Then
        ReDim Preserve IBlackRollBoard(UBound(IBlackRollBoard) + 1) As Integer
    End If
    
    IBlackRollBoard(UBound(IBlackRollBoard)) = IValue
    
    If BBlack Then
        If IValue = INone Then
            'If BEven Or BHigh Then
            'Else
                If IWin > 0 Then
                    IWin = IWin - 1
                Else
                    ILoss = ILoss + 1
                End If
            'End If
        Else
            If BBlackPattern Then
                If IValue = ITick Then
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
            Else
                If IValue = ICross Then
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
            End If
        End If
    End If
End Sub

Public Sub SetEvenRollBoard(ByVal IValue As Integer)
    If Not (IEvenRollBoard(UBound(IEvenRollBoard)) = IBlank) Then
        ReDim Preserve IEvenRollBoard(UBound(IEvenRollBoard) + 1) As Integer
    End If
    
    IEvenRollBoard(UBound(IEvenRollBoard)) = IValue
    
    If BEven Then
        If IValue = INone Then
            'If BBlack Or BHigh Then
            'Else
                If IWin > 0 Then
                    IWin = IWin - 1
                Else
                    ILoss = ILoss + 1
                End If
            'End If
        Else
            If BEvenPattern Then
                If IValue = ITick Then
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
            Else
                If IValue = ICross Then
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
            End If
        End If
    End If
End Sub

Public Sub SetHighRollBoard(ByVal IValue As Integer)
    If Not (IHighRollBoard(UBound(IHighRollBoard)) = IBlank) Then
        ReDim Preserve IHighRollBoard(UBound(IHighRollBoard) + 1) As Integer
    End If
    
    IHighRollBoard(UBound(IHighRollBoard)) = IValue
    
    If BHigh Then
        If IValue = INone Then
            'If BBlack Or BEven Then
            'Else
                If IWin > 0 Then
                    IWin = IWin - 1
                Else
                    ILoss = ILoss + 1
                End If
            'End If
        Else
            If BHighPattern Then
                If IValue = ITick Then
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
            Else
                If IValue = ICross Then
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
            End If
        End If
    End If
End Sub

Public Sub SetHistory(ByVal IValue As Integer, Optional ByVal BSave As Boolean = True)
    On Local Error GoTo ErrHandler
    
    If UBound(IRollHTS) >= IHistoryMax Then
    Else
        If Not (IRollHTS(UBound(IRollHTS)) = IBlank) Then
            ReDim Preserve IRollHTS(UBound(IRollHTS) + 1) As Integer
        End If
        
        IRollHTS(UBound(IRollHTS)) = IValue
        
        If BSave Then
            Dim lRegistry As Long
            Dim lRegKey As Long
            
            lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, lRegKey)
            If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY)
            lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, CStr(UBound(IRollHTS) + 1), CStr(IRollHTS(UBound(IRollHTS))))
            lRegistry = mdAPI.CloseRegistry(lRegKey)
        End If
    End If
    
    Exit Sub
ErrHandler:
End Sub

Public Sub SetWin(Optional ByVal IValue As Integer = 0)
    IWin = IValue
End Sub

Public Sub SetLoss(Optional ByVal IValue As Integer = 0)
    ILoss = IValue
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

Public Sub SetBlackPattern(ByVal IValue As Integer)
    If UBound(IBlackRollPattern) = 0 And (IBlackRollPattern(UBound(IBlackRollPattern)) = IBlank) Then
        IBlackRollPattern(UBound(IBlackRollPattern)) = IValue
    Else
        ReDim Preserve IBlackRollPattern(UBound(IBlackRollPattern) + 1) As Integer
        
        IBlackRollPattern(UBound(IBlackRollPattern)) = IValue
    End If
End Sub

Public Sub SetEvenPattern(ByVal IValue As Integer)
    If UBound(IEvenRollPattern) = 0 And (IEvenRollPattern(UBound(IEvenRollPattern)) = IBlank) Then
        IEvenRollPattern(UBound(IEvenRollPattern)) = IValue
    Else
        ReDim Preserve IEvenRollPattern(UBound(IEvenRollPattern) + 1) As Integer
        
        IEvenRollPattern(UBound(IEvenRollPattern)) = IValue
    End If
End Sub

Public Sub SetHighPattern(ByVal IValue As Integer)
    If UBound(IHighRollPattern) = 0 And (IHighRollPattern(UBound(IHighRollPattern)) = IBlank) Then
        IHighRollPattern(UBound(IHighRollPattern)) = IValue
    Else
        ReDim Preserve IHighRollPattern(UBound(IHighRollPattern) + 1) As Integer
        
        IHighRollPattern(UBound(IHighRollPattern)) = IValue
    End If
End Sub

Public Sub SetBlackBox(ByVal IType As WinLossBox)
    If IBlackBox(4) >= IBoxBaseDec Then Exit Sub
    
    If IType = WIN Then
        If IBlackBox(1) > 0 Then
            IBlackBox(1) = IBlackBox(1) - 1
        Else
            IBlackBox(0) = IBlackBox(0) + 1
        End If
        
        If IRedBox(9) = ITick Then
            If IRedBox(5) > 0 Then
                IRedBox(5) = IRedBox(5) - 1
            Else
                IRedBox(6) = IRedBox(6) + 1
            End If
        End If
        
        If IBlackBox(0) > IBlackBox(2) Then IBlackBox(2) = IBlackBox(0)
        If IRedBox(5) > IRedBox(7) Then IRedBox(7) = IRedBox(5)
    ElseIf IType = LOSS Then
        If IBlackBox(0) > 0 Then
            IBlackBox(0) = IBlackBox(0) - 1
        Else
            IBlackBox(1) = IBlackBox(1) + 1
        End If
        
        If IRedBox(9) = ITick Then
            If IRedBox(6) > 0 Then
                IRedBox(6) = IRedBox(6) - 1
            Else
                IRedBox(5) = IRedBox(5) + 1
            End If
        End If
        
        If IBlackBox(1) > IBlackBox(3) Then IBlackBox(3) = IBlackBox(1)
        If IRedBox(6) > IRedBox(8) Then IRedBox(8) = IRedBox(6)
    End If
    
    If (IBlackBox(0) = 0) And (IBlackBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IBlackBox)
            IBlackBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IRedBox)
            IRedBox(ICounter) = 0
        Next ICounter
    Else
        IBlackBox(4) = IBlackBox(4) + 1
    End If
    
    If IBlackBox(0) = IBoxWin Then
        If Not (IRedBox(9) = ITick) Then IRedBox(9) = ITick
    End If
End Sub

Public Sub SetRBlackBox(ByVal IType As WinLossBox)
    If IType = WIN Then
        If IBlackBox(0) > 0 Then
            IBlackBox(0) = IBlackBox(0) - 1
        Else
            IBlackBox(1) = IBlackBox(1) + 1
        End If
        
        If IRedBox(9) = ITick Then
            If IRedBox(6) > 0 Then
                IRedBox(6) = IRedBox(6) - 1
            Else
                IRedBox(5) = IRedBox(5) + 1
            End If
        End If
    ElseIf IType = LOSS Then
        If IBlackBox(1) > 0 Then
            IBlackBox(1) = IBlackBox(1) - 1
        Else
            IBlackBox(0) = IBlackBox(0) + 1
        End If
        
        If IRedBox(9) = ITick Then
            If IRedBox(5) > 0 Then
                IRedBox(5) = IRedBox(5) - 1
            Else
                IRedBox(6) = IRedBox(6) + 1
            End If
        End If
    End If
    
    If (IBlackBox(0) = 0) And (IBlackBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IBlackBox)
            IBlackBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IRedBox)
            IRedBox(ICounter) = 0
        Next ICounter
    Else
        IBlackBox(4) = IBlackBox(4) - 1
    End If
End Sub

Public Sub SetRedBox(ByVal IType As WinLossBox)
    If IRedBox(4) >= IBoxBaseDec Then Exit Sub
    
    If IType = WIN Then
        If IRedBox(1) > 0 Then
            IRedBox(1) = IRedBox(1) - 1
        Else
            IRedBox(0) = IRedBox(0) + 1
        End If
        
        If IBlackBox(9) = ITick Then
            If IBlackBox(5) > 0 Then
                IBlackBox(5) = IBlackBox(5) - 1
            Else
                IBlackBox(6) = IBlackBox(6) + 1
            End If
        End If
        
        If IRedBox(0) > IRedBox(2) Then IRedBox(2) = IRedBox(0)
        If IBlackBox(5) > IBlackBox(7) Then IBlackBox(7) = IBlackBox(5)
    ElseIf IType = LOSS Then
        If IRedBox(0) > 0 Then
            IRedBox(0) = IRedBox(0) - 1
        Else
            IRedBox(1) = IRedBox(1) + 1
        End If
        
        If IBlackBox(9) = ITick Then
            If IBlackBox(6) > 0 Then
                IBlackBox(6) = IBlackBox(6) - 1
            Else
                IBlackBox(5) = IBlackBox(5) + 1
            End If
        End If
        
        If IRedBox(1) > IRedBox(3) Then IRedBox(3) = IRedBox(1)
        If IBlackBox(6) > IBlackBox(8) Then IBlackBox(8) = IBlackBox(6)
    End If
    
    If (IRedBox(0) = 0) And (IRedBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IRedBox)
            IRedBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IBlackBox)
            IBlackBox(ICounter) = 0
        Next ICounter
    Else
        IRedBox(4) = IRedBox(4) + 1
    End If
    
    If IRedBox(0) = IBoxWin Then
        If Not (IBlackBox(9) = ITick) Then IBlackBox(9) = ITick
    End If
End Sub

Public Sub SetRRedBox(ByVal IType As WinLossBox)
    If IType = WIN Then
        If IRedBox(0) > 0 Then
            IRedBox(0) = IRedBox(0) - 1
        Else
            IRedBox(1) = IRedBox(1) + 1
        End If
        
        If IBlackBox(9) = ITick Then
            If IBlackBox(6) > 0 Then
                IBlackBox(6) = IBlackBox(6) - 1
            Else
                IBlackBox(5) = IBlackBox(5) + 1
            End If
        End If
    ElseIf IType = LOSS Then
        If IRedBox(1) > 0 Then
            IRedBox(1) = IRedBox(1) - 1
        Else
            IRedBox(0) = IRedBox(0) + 1
        End If
        
        If IBlackBox(9) = ITick Then
            If IBlackBox(5) > 0 Then
                IBlackBox(5) = IBlackBox(5) - 1
            Else
                IBlackBox(6) = IBlackBox(6) + 1
            End If
        End If
    End If
    
    If (IRedBox(0) = 0) And (IRedBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IRedBox)
            IRedBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IBlackBox)
            IBlackBox(ICounter) = 0
        Next ICounter
    Else
        IRedBox(4) = IRedBox(4) - 1
    End If
End Sub

Public Sub SetEvenBox(ByVal IType As WinLossBox)
    If IEvenBox(4) >= IBoxBaseDec Then Exit Sub
    
    If IType = WIN Then
        If IEvenBox(1) > 0 Then
            IEvenBox(1) = IEvenBox(1) - 1
        Else
            IEvenBox(0) = IEvenBox(0) + 1
        End If
        
        If IOddBox(9) = ITick Then
            If IOddBox(5) > 0 Then
                IOddBox(5) = IOddBox(5) - 1
            Else
                IOddBox(6) = IOddBox(6) + 1
            End If
        End If
        
        If IEvenBox(0) > IEvenBox(2) Then IEvenBox(2) = IEvenBox(0)
        If IOddBox(5) > IOddBox(7) Then IOddBox(7) = IOddBox(5)
    ElseIf IType = LOSS Then
        If IEvenBox(0) > 0 Then
            IEvenBox(0) = IEvenBox(0) - 1
        Else
            IEvenBox(1) = IEvenBox(1) + 1
        End If
        
        If IOddBox(9) = ITick Then
            If IOddBox(6) > 0 Then
                IOddBox(6) = IOddBox(6) - 1
            Else
                IOddBox(5) = IOddBox(5) + 1
            End If
        End If
        
        If IEvenBox(1) > IEvenBox(3) Then IEvenBox(3) = IEvenBox(1)
        If IOddBox(6) > IOddBox(8) Then IOddBox(8) = IOddBox(6)
    End If
    
    If (IEvenBox(0) = 0) And (IEvenBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IEvenBox)
            IEvenBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IOddBox)
            IOddBox(ICounter) = 0
        Next ICounter
    Else
        IEvenBox(4) = IEvenBox(4) + 1
    End If
    
    If IEvenBox(0) = IBoxWin Then
        If Not (IOddBox(9) = ITick) Then IOddBox(9) = ITick
    End If
End Sub

Public Sub SetREvenBox(ByVal IType As WinLossBox)
    If IType = WIN Then
        If IEvenBox(0) > 0 Then
            IEvenBox(0) = IEvenBox(0) - 1
        Else
            IEvenBox(1) = IEvenBox(1) + 1
        End If
        
        If IOddBox(9) = ITick Then
            If IOddBox(6) > 0 Then
                IOddBox(6) = IOddBox(6) - 1
            Else
                IOddBox(5) = IOddBox(5) + 1
            End If
        End If
    ElseIf IType = LOSS Then
        If IEvenBox(1) > 0 Then
            IEvenBox(1) = IEvenBox(1) - 1
        Else
            IEvenBox(0) = IEvenBox(0) + 1
        End If
        
        If IOddBox(9) = ITick Then
            If IOddBox(5) > 0 Then
                IOddBox(5) = IOddBox(5) - 1
            Else
                IOddBox(6) = IOddBox(6) + 1
            End If
        End If
    End If
    
    If (IEvenBox(0) = 0) And (IEvenBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IEvenBox)
            IEvenBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IOddBox)
            IOddBox(ICounter) = 0
        Next ICounter
    Else
        IEvenBox(4) = IEvenBox(4) - 1
    End If
End Sub

Public Sub SetOddBox(ByVal IType As WinLossBox)
    If IOddBox(4) >= IBoxBaseDec Then Exit Sub
    
    If IType = WIN Then
        If IOddBox(1) > 0 Then
            IOddBox(1) = IOddBox(1) - 1
        Else
            IOddBox(0) = IOddBox(0) + 1
        End If
        
        If IEvenBox(9) = ITick Then
            If IEvenBox(5) > 0 Then
                IEvenBox(5) = IEvenBox(5) - 1
            Else
                IEvenBox(6) = IEvenBox(6) + 1
            End If
        End If
        
        If IOddBox(0) > IOddBox(2) Then IOddBox(2) = IOddBox(0)
        If IEvenBox(5) > IEvenBox(7) Then IEvenBox(7) = IEvenBox(5)
    ElseIf IType = LOSS Then
        If IOddBox(0) > 0 Then
            IOddBox(0) = IOddBox(0) - 1
        Else
            IOddBox(1) = IOddBox(1) + 1
        End If
        
        If IEvenBox(9) = ITick Then
            If IEvenBox(6) > 0 Then
                IEvenBox(6) = IEvenBox(6) - 1
            Else
                IEvenBox(5) = IEvenBox(5) + 1
            End If
        End If
        
        If IOddBox(1) > IOddBox(3) Then IOddBox(3) = IOddBox(1)
        If IEvenBox(6) > IEvenBox(8) Then IEvenBox(8) = IEvenBox(6)
    End If
    
    If (IOddBox(0) = 0) And (IOddBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IOddBox)
            IOddBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IEvenBox)
            IEvenBox(ICounter) = 0
        Next ICounter
    Else
        IOddBox(4) = IOddBox(4) + 1
    End If
    
    If IOddBox(0) = IBoxWin Then
        If Not (IEvenBox(9) = ITick) Then IEvenBox(9) = ITick
    End If
End Sub

Public Sub SetROddBox(ByVal IType As WinLossBox)
    If IType = WIN Then
        If IOddBox(0) > 0 Then
            IOddBox(0) = IOddBox(0) - 1
        Else
            IOddBox(1) = IOddBox(1) + 1
        End If
        
        If IEvenBox(9) = ITick Then
            If IEvenBox(6) > 0 Then
                IEvenBox(6) = IEvenBox(6) - 1
            Else
                IEvenBox(5) = IEvenBox(5) + 1
            End If
        End If
    ElseIf IType = LOSS Then
        If IOddBox(1) > 0 Then
            IOddBox(1) = IOddBox(1) - 1
        Else
            IOddBox(0) = IOddBox(0) + 1
        End If
        
        If IEvenBox(9) = ITick Then
            If IEvenBox(5) > 0 Then
                IEvenBox(5) = IEvenBox(5) - 1
            Else
                IEvenBox(6) = IEvenBox(6) + 1
            End If
        End If
    End If
    
    If (IOddBox(0) = 0) And (IOddBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IOddBox)
            IOddBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IEvenBox)
            IEvenBox(ICounter) = 0
        Next ICounter
    Else
        IOddBox(4) = IOddBox(4) - 1
    End If
End Sub

Public Sub SetHighBox(ByVal IType As WinLossBox)
    If IHighBox(4) >= IBoxBaseDec Then Exit Sub
    
    If IType = WIN Then
        If IHighBox(1) > 0 Then
            IHighBox(1) = IHighBox(1) - 1
        Else
            IHighBox(0) = IHighBox(0) + 1
        End If
        
        If ILowBox(9) = ITick Then
            If ILowBox(5) > 0 Then
                ILowBox(5) = ILowBox(5) - 1
            Else
                ILowBox(6) = ILowBox(6) + 1
            End If
        End If
        
        If IHighBox(0) > IHighBox(2) Then IHighBox(2) = IHighBox(0)
        If ILowBox(5) > ILowBox(7) Then ILowBox(7) = ILowBox(5)
    ElseIf IType = LOSS Then
        If IHighBox(0) > 0 Then
            IHighBox(0) = IHighBox(0) - 1
        Else
            IHighBox(1) = IHighBox(1) + 1
        End If
        
        If ILowBox(9) = ITick Then
            If ILowBox(6) > 0 Then
                ILowBox(6) = ILowBox(6) - 1
            Else
                ILowBox(5) = ILowBox(5) + 1
            End If
        End If
        
        If IHighBox(1) > IHighBox(3) Then IHighBox(3) = IHighBox(1)
        If ILowBox(6) > ILowBox(8) Then ILowBox(8) = ILowBox(6)
    End If
    
    If (IHighBox(0) = 0) And (IHighBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IHighBox)
            IHighBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(ILowBox)
            ILowBox(ICounter) = 0
        Next ICounter
    Else
        IHighBox(4) = IHighBox(4) + 1
    End If
    
    If IHighBox(0) = IBoxWin Then
        If Not (ILowBox(9) = ITick) Then ILowBox(9) = ITick
    End If
End Sub

Public Sub SetRHighBox(ByVal IType As WinLossBox)
    If IType = WIN Then
        If IHighBox(0) > 0 Then
            IHighBox(0) = IHighBox(0) - 1
        Else
            IHighBox(1) = IHighBox(1) + 1
        End If
        
        If ILowBox(9) = ITick Then
            If ILowBox(6) > 0 Then
                ILowBox(6) = ILowBox(6) - 1
            Else
                ILowBox(5) = ILowBox(5) + 1
            End If
        End If
    ElseIf IType = LOSS Then
        If IHighBox(1) > 0 Then
            IHighBox(1) = IHighBox(1) - 1
        Else
            IHighBox(0) = IHighBox(0) + 1
        End If
        
        If ILowBox(9) = ITick Then
            If ILowBox(5) > 0 Then
                ILowBox(5) = ILowBox(5) - 1
            Else
                ILowBox(6) = ILowBox(6) + 1
            End If
        End If
    End If
    
    If (IHighBox(0) = 0) And (IHighBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(IHighBox)
            IHighBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(ILowBox)
            ILowBox(ICounter) = 0
        Next ICounter
    Else
        IHighBox(4) = IHighBox(4) - 1
    End If
End Sub

Public Sub SetLowBox(ByVal IType As WinLossBox)
    If ILowBox(4) >= IBoxBaseDec Then Exit Sub
    
    If IType = WIN Then
        If ILowBox(1) > 0 Then
            ILowBox(1) = ILowBox(1) - 1
        Else
            ILowBox(0) = ILowBox(0) + 1
        End If
        
        If IHighBox(9) = ITick Then
            If IHighBox(5) > 0 Then
                IHighBox(5) = IHighBox(5) - 1
            Else
                IHighBox(6) = IHighBox(6) + 1
            End If
        End If
        
        If ILowBox(0) > ILowBox(2) Then ILowBox(2) = ILowBox(0)
        If IHighBox(5) > IHighBox(7) Then IHighBox(7) = IHighBox(5)
    ElseIf IType = LOSS Then
        If ILowBox(0) > 0 Then
            ILowBox(0) = ILowBox(0) - 1
        Else
            ILowBox(1) = ILowBox(1) + 1
        End If
        
        If IHighBox(9) = ITick Then
            If IHighBox(6) > 0 Then
                IHighBox(6) = IHighBox(6) - 1
            Else
                IHighBox(5) = IHighBox(5) + 1
            End If
        End If
        
        If ILowBox(1) > ILowBox(3) Then ILowBox(3) = ILowBox(1)
        If IHighBox(6) > IHighBox(8) Then IHighBox(8) = IHighBox(6)
    End If
    
    If (ILowBox(0) = 0) And (ILowBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(ILowBox)
            ILowBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IHighBox)
            IHighBox(ICounter) = 0
        Next ICounter
    Else
        ILowBox(4) = ILowBox(4) + 1
    End If
    
    If ILowBox(0) = IBoxWin Then
        If Not (IHighBox(9) = ITick) Then IHighBox(9) = ITick
    End If
End Sub

Public Sub SetRLowBox(ByVal IType As WinLossBox)
    If IType = WIN Then
        If ILowBox(0) > 0 Then
            ILowBox(0) = ILowBox(0) - 1
        Else
            ILowBox(1) = ILowBox(1) + 1
        End If
        
        If IHighBox(9) = ITick Then
            If IHighBox(6) > 0 Then
                IHighBox(6) = IHighBox(6) - 1
            Else
                IHighBox(5) = IHighBox(5) + 1
            End If
        End If
    ElseIf IType = LOSS Then
        If ILowBox(1) > 0 Then
            ILowBox(1) = ILowBox(1) - 1
        Else
            ILowBox(0) = ILowBox(0) + 1
        End If
        
        If IHighBox(9) = ITick Then
            If IHighBox(5) > 0 Then
                IHighBox(5) = IHighBox(5) - 1
            Else
                IHighBox(6) = IHighBox(6) + 1
            End If
        End If
    End If
    
    If (ILowBox(0) = 0) And (ILowBox(1) = 0) Then
        Dim ICounter As Integer
        
        For ICounter = 2 To UBound(ILowBox)
            ILowBox(ICounter) = 0
        Next ICounter
        
        For ICounter = 5 To UBound(IHighBox)
            IHighBox(ICounter) = 0
        Next ICounter
    Else
        ILowBox(4) = ILowBox(4) - 1
    End If
End Sub

Public Function GetBlackBox(Optional ByVal IElement As Integer = 0) As Integer
    GetBlackBox = IBlackBox(IElement)
End Function

Public Function GetRedBox(Optional ByVal IElement As Integer = 0) As Integer
    GetRedBox = IRedBox(IElement)
End Function

Public Function GetEvenBox(Optional ByVal IElement As Integer = 0) As Integer
    GetEvenBox = IEvenBox(IElement)
End Function

Public Function GetOddBox(Optional ByVal IElement As Integer = 0) As Integer
    GetOddBox = IOddBox(IElement)
End Function

Public Function GetHighBox(Optional ByVal IElement As Integer = 0) As Integer
    GetHighBox = IHighBox(IElement)
End Function

Public Function GetLowBox(Optional ByVal IElement As Integer = 0) As Integer
    GetLowBox = ILowBox(IElement)
End Function

Public Sub InitRollBoard()
    ReDim INumberRollBoard(0) As Integer
    ReDim IBlackRollBoard(0) As Integer
    ReDim IEvenRollBoard(0) As Integer
    ReDim IHighRollBoard(0) As Integer
    
    INumberRollBoard(0) = IBlank
    IBlackRollBoard(0) = IBlank
    IEvenRollBoard(0) = IBlank
    IHighRollBoard(0) = IBlank
End Sub

Public Sub InitRollPattern()
    ReDim IBlackRollPattern(0) As Integer
    ReDim IEvenRollPattern(0) As Integer
    ReDim IHighRollPattern(0) As Integer
    
    IBlackRollPattern(0) = IBlank
    IEvenRollPattern(0) = IBlank
    IHighRollPattern(0) = IBlank
    
    IWin = 0
    ILoss = 0
    
    BBlack = False
    BEven = False
    BHigh = False
End Sub

Public Sub InitRollHistory()
    ITable = 1
    
    ReDim IRollHT1(INumberMax) As Integer
    ReDim IRollHT2(INumberMax) As Integer
    
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_TABLE, lRegKey)

    If lRegistry = 0 Then
        Dim SValueC() As String
        
        Dim SValueHT1 As String
        Dim SValueHT2 As String
        Dim SValue As String
        Dim STable As String
        
        Dim ICounter As Integer
        
        For ICounter = 0 To INumberMax
            lRegistry = mdAPI.ReadValueRegistry(lRegKey, CStr(ICounter) & "-1", lType, SValueHT1, lSize)
            lRegistry = mdAPI.ReadValueRegistry(lRegKey, CStr(ICounter) & "-2", lType, SValueHT2, lSize)
            SValueHT1 = mdAPI.ReplaceRegistry(SValueHT1)
            SValueHT2 = mdAPI.ReplaceRegistry(SValueHT2)
            
            If IsNumeric(SValueHT1) Then IRollHT1(ICounter) = CInt(SValueHT1)
            If IsNumeric(SValueHT2) Then IRollHT2(ICounter) = CInt(SValueHT2)
        Next ICounter
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, "Table", lType, STable, lSize)
        STable = mdAPI.ReplaceRegistry(STable)
        
        If IsNumeric(STable) Then ITable = CInt(STable)
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, "CNT - 1", lType, SValue, lSize)
        SValue = mdAPI.ReplaceRegistry(SValue)
        
        If Not (Trim(SValue) = "") Then
            SValueC = Split(SValue, "|")
            
            For ICounter = LBound(SValueC) To UBound(SValueC)
                If IsNumeric(SValueC(ICounter)) Then IRollCHT1(ICounter) = CInt(SValueC(ICounter))
            Next ICounter
        End If
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, "CNT - 2", lType, SValue, lSize)
        SValue = mdAPI.ReplaceRegistry(SValue)
        
        If Not (Trim(SValue) = "") Then
            SValueC = Split(SValue, "|")
            
            For ICounter = LBound(SValueC) To UBound(SValueC)
                If IsNumeric(SValueC(ICounter)) Then IRollCHT2(ICounter) = CInt(SValueC(ICounter))
            Next ICounter
        End If
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    
    If lRegistry = 0 Then
        Dim SValueRoll As String
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, "ROLL DEC", lType, SValueRoll, lSize)
        SValueRoll = mdAPI.ReplaceRegistry(SValueRoll)
        
        If IsNumeric(SValueRoll) Then
            ICounterRollDec = CInt(SValueRoll)
        Else
            ICounterRollDec = IRollBaseDec
        End If
        
        For ICounter = 1 To (IRollBaseDec - ICounterRollDec + 1)
            lRegistry = mdAPI.ReadValueRegistry(lRegKey, "ROW - " & CStr(ICounter), lType, SValueRoll, lSize)
            SValueRoll = mdAPI.ReplaceRegistry(SValueRoll)
            
            If IsNumeric(SValueRoll) Then
                If CInt(SValueRoll) = IBlank Then
                Else
                    SetNumberRollBoard CInt(SValueRoll), False
                    
                    If mdApp.LColorPattern(CInt(SValueRoll)) = mdApp.LRed Then
                        mdApp.SetBlackRollBoard mdApp.ICross
                        mdApp.SetBlackPattern mdApp.ICross
                    ElseIf LColorPattern(CInt(SValueRoll)) = mdApp.LBlack Then
                        mdApp.SetBlackRollBoard mdApp.ITick
                        mdApp.SetBlackPattern mdApp.ITick
                    ElseIf LColorPattern(CInt(SValueRoll)) = mdApp.LGreen Then
                        mdApp.SetBlackRollBoard mdApp.INone
                        mdApp.SetBlackPattern mdApp.INone
                    End If
                    
                    If CInt(SValueRoll) = 0 Then
                        mdApp.SetEvenRollBoard mdApp.INone
                        mdApp.SetEvenPattern mdApp.INone
                    ElseIf CInt(SValueRoll) Mod 2 = 0 Then
                        mdApp.SetEvenRollBoard mdApp.ITick
                        mdApp.SetEvenPattern mdApp.ITick
                    Else
                        mdApp.SetEvenRollBoard mdApp.ICross
                        mdApp.SetEvenPattern mdApp.ICross
                    End If
                    
                    If CInt(SValueRoll) = 0 Then
                        mdApp.SetHighRollBoard mdApp.INone
                        mdApp.SetHighPattern mdApp.INone
                    ElseIf CInt(SValueRoll) >= 19 Then
                        mdApp.SetHighRollBoard mdApp.ITick
                        mdApp.SetHighPattern mdApp.ITick
                    Else
                        mdApp.SetHighRollBoard mdApp.ICross
                        mdApp.SetHighPattern mdApp.ICross
                    End If
                End If
            End If
        Next ICounter
    Else
        ICounterRollDec = IRollBaseDec
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub InitHistory()
    ReDim IRollHTS(0) As Integer
    IRollHTS(0) = IBlank
    
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, lRegKey)

    If lRegistry = 0 Then
        Dim ICounter As Integer
        
        Dim SValueHTS As String
        
        For ICounter = 1 To IHistoryMax
            lRegistry = mdAPI.ReadValueRegistry(lRegKey, CStr(ICounter), lType, SValueHTS, lSize)
            SValueHTS = mdAPI.ReplaceRegistry(SValueHTS)
            
            If IsNumeric(SValueHTS) Then SetHistory CInt(SValueHTS), False
        Next ICounter
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub InitPatternHistory()
    ReDim IRollBPatternHTS(UBound(IRollHTS)) As Integer
    ReDim IRollEPatternHTS(UBound(IRollHTS)) As Integer
    ReDim IRollHPatternHTS(UBound(IRollHTS)) As Integer
    
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_HISTORY, lRegKey)

    If lRegistry = 0 Then
        Dim SValueHTS As String
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SBlackText, lType, SValueHTS, lSize)
        SValueHTS = mdAPI.ReplaceRegistry(SValueHTS)
        
        SetPatternHistory IRollBPatternHTS, SValueHTS
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SEvenText, lType, SValueHTS, lSize)
        SValueHTS = mdAPI.ReplaceRegistry(SValueHTS)
        
        SetPatternHistory IRollEPatternHTS, SValueHTS
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SHighText, lType, SValueHTS, lSize)
        SValueHTS = mdAPI.ReplaceRegistry(SValueHTS)
        
        SetPatternHistory IRollHPatternHTS, SValueHTS
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub InitBox()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX, lRegKey)

    If lRegistry = 0 Then
        Dim ICounter As Integer
        Dim SValueBox As String
        Dim SValue() As String
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SBlackText, lType, SValueBox, lSize)
        SValueBox = mdAPI.ReplaceRegistry(SValueBox)
        
        SValue = Split(SValueBox, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                IBlackBox(ICounter) = SValue(ICounter)
            Else
                IBlackBox(ICounter) = INone
            End If
        Next ICounter
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SRedText, lType, SValueBox, lSize)
        SValueBox = mdAPI.ReplaceRegistry(SValueBox)
        
        SValue = Split(SValueBox, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                IRedBox(ICounter) = SValue(ICounter)
            Else
                IRedBox(ICounter) = INone
            End If
        Next ICounter
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SEvenText, lType, SValueBox, lSize)
        SValueBox = mdAPI.ReplaceRegistry(SValueBox)
        
        SValue = Split(SValueBox, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                IEvenBox(ICounter) = SValue(ICounter)
            Else
                IEvenBox(ICounter) = INone
            End If
        Next ICounter
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SOddText, lType, SValueBox, lSize)
        SValueBox = mdAPI.ReplaceRegistry(SValueBox)
        
        SValue = Split(SValueBox, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                IOddBox(ICounter) = SValue(ICounter)
            Else
                IOddBox(ICounter) = INone
            End If
        Next ICounter
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SHighText, lType, SValueBox, lSize)
        SValueBox = mdAPI.ReplaceRegistry(SValueBox)
        
        SValue = Split(SValueBox, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                IHighBox(ICounter) = SValue(ICounter)
            Else
                IHighBox(ICounter) = INone
            End If
        Next ICounter
        
        lRegistry = mdAPI.ReadValueRegistry(lRegKey, SLowText, lType, SValueBox, lSize)
        SValueBox = mdAPI.ReplaceRegistry(SValueBox)
        
        SValue = Split(SValueBox, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                ILowBox(ICounter) = SValue(ICounter)
            Else
                ILowBox(ICounter) = INone
            End If
        Next ICounter
    Else
        lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_BOX)
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Private Sub SetPatternHistory(ByRef IRollPatternHTS() As Integer, ByVal SValueHTS As String)
    If Not (Trim(SValueHTS) = "") Then
        Dim SValue() As String
        Dim ICounter As Integer
        
        SValue = Split(SValueHTS, "|")
        
        For ICounter = LBound(SValue) To UBound(SValue)
            If IsNumeric(SValue(ICounter)) Then
                IRollPatternHTS(ICounter) = SValue(ICounter)
            Else
                IRollBPatternHTS(ICounter) = INone
            End If
        Next ICounter
    End If
End Sub

Public Function CheckPattern(Optional ByVal BAlert As Boolean = True) As String
    Dim SBlack As String
    Dim SEven As String
    Dim SHigh As String
    
    SBlack = CheckProbPattern(IBlackRollPattern, IBlackType)
    SEven = CheckProbPattern(IEvenRollPattern, IEvenType)
    SHigh = CheckProbPattern(IHighRollPattern, IHighType)
    
    Dim SValue As String
    
    SValue = SBlack
    SValue = SValue & vbCrLf
    SValue = SValue & SEven
    SValue = SValue & vbCrLf
    SValue = SValue & SHigh
    
    If BAlert Then
        If Trim(SBlack) = "" And Trim(SEven) = "" And Trim(SHigh) = "" Then
            mdAPI.Beep 300, 50
        Else
            mdAPI.Beep 700, 300
        End If
    End If

    CheckPattern = SValue
End Function

Public Function CheckZero() As Integer
    Dim ICounter As Integer
    Dim ICheck As Integer
    
    CheckZero = 0
    ICheck = 0
    
    For ICounter = UBound(INumberRollBoard) To LBound(INumberRollBoard) Step -1
        If ICheck = 6 Then Exit For
        
        If INumberRollBoard(ICounter) = INone Then
            CheckZero = CheckZero + 1
        End If
        
        ICheck = ICheck + 1
    Next ICounter
End Function

Private Function CheckProbPattern(ByRef IPattern() As Integer, ByVal IType As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False) As String
    Dim ITemp() As Integer
    Dim ICounter As Integer
    Dim BDoubleTick As Boolean
    Dim BDoubleCross As Boolean
    Dim BDoubleNone As Boolean
    Dim SOutput As String
    Dim SValue As String
    
    CheckProbPattern = ""
    SOutput = ""
    
    If IType = IBlackType Then
        If BScore Then IBlackScore = IBlank
        If BFocus Then IBlackFocus = 0
        
        IBOutcome = 0
        BBlack = False
    ElseIf IType = IEvenType Then
        If BScore Then IEvenScore = IBlank
        If BFocus Then IEvenFocus = 0
    
        IEOutcome = 0
        BEven = False
    ElseIf IType = IHighType Then
        If BScore Then IHighScore = IBlank
        If BFocus Then IHighFocus = 0
        
        IHOutcome = 0
        BHigh = False
    End If
    
    For ICounter = UBound(IPattern) To LBound(IPattern) + 5 Step -1
        SValue = ""
        BDoubleTick = False
        BDoubleCross = False
        BDoubleNone = False
        
        If IPattern(ICounter) = IPattern(ICounter - 1) Then
            If IPattern(ICounter) = ITick Then
                BDoubleTick = True
            ElseIf IPattern(ICounter) = ICross Then
                BDoubleCross = True
            ElseIf IPattern(ICounter) = INone Then
                BDoubleNone = True
            End If
        
            If IPattern(ICounter - 1) = IPattern(ICounter - 2) Then
                If BDoubleNone Then
                    SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                Else
                    If BWinLoss Then SetWinLoss
                    
                    Exit For
                End If
            ElseIf IPattern(ICounter - 2) = IPattern(ICounter - 3) Then
                If IPattern(ICounter - 3) = IPattern(ICounter - 4) Then
                    If BWinLoss Then SetWinLoss
                    
                    Exit For
                ElseIf IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                    If BDoubleNone Then
                        SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    Else
                        If BWinLoss Then SetWinLoss
                        
                        Exit For
                    End If
                Else
                    If BDoubleTick Then
                        If IPattern(ICounter - 4) = INone Then
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText2)
                            
                            SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                        End If
                    ElseIf BDoubleCross Then
                        If IPattern(ICounter - 4) = INone Then '2,1,0,2,2,1,1
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText1)
                            
                            SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                        End If
                    ElseIf BDoubleNone Then
                        If (IPattern(ICounter - 4) = INone) Or (IPattern(ICounter - 5) = INone) Then
                            SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        End If
                    End If
                End If
            ElseIf IPattern(ICounter - 3) = IPattern(ICounter - 4) Then
                If BDoubleNone Then
                    If IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                        If BWinLoss Then SetWinLoss
                        
                        Exit For
                    Else
                        If IPattern(ICounter - 5) = INone Then
                            SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        End If
                    End If
                Else
                    If IPattern(ICounter - 2) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    Else
                        If BWinLoss Then SetWinLoss
                        
                        Exit For
                    End If
                End If
            ElseIf IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                If BDoubleNone Then
                    SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                Else
                    If IPattern(ICounter - 2) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    ElseIf IPattern(ICounter - 3) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    ElseIf IPattern(ICounter) = ITick Then
                        If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText2)
                        
                        SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                    ElseIf IPattern(ICounter) = ICross Then
                        If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText1)
                        
                        SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                    End If
                End If
            Else
                If BDoubleNone Then
                    If IPattern(ICounter - 5) = INone Then
                        SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    End If
                Else
                    If IPattern(ICounter - 2) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    ElseIf IPattern(ICounter - 3) = INone Then '1,2,0,1,2,2
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    ElseIf IPattern(ICounter - 4) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        
                        If Trim(SValue) = "" Then
                            If BWinLoss Then SetWinLoss
                            
                            Exit For
                        End If
                    ElseIf IPattern(ICounter - 5) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    Else
                        If BWinLoss Then SetWinLoss
                        
                        Exit For
                    End If
                End If
            End If
        Else
            If IPattern(ICounter - 1) = IPattern(ICounter - 2) Then
                If IPattern(ICounter - 1) = INone Then
                    BDoubleNone = True
                End If
            
                If IPattern(ICounter - 2) = IPattern(ICounter - 3) Then
                    If BDoubleNone Then
                        SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    Else
                        If BWinLoss Then SetWinLoss
                        
                        Exit For
                    End If
                ElseIf IPattern(ICounter - 3) = IPattern(ICounter - 4) Then
                    If IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                        If IPattern(ICounter - 3) = INone Then
                            If IPattern(ICounter) = IPattern(ICounter - 3) Then
                                If BWinLoss Then SetWinLoss
                                
                                Exit For
                            Else
                                SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            End If
                        Else
                            If BWinLoss Then SetWinLoss
                            
                            Exit For
                        End If
                    Else
                        If IPattern(ICounter - 3) = INone Then
                            If BDoubleNone Then
                                If BWinLoss Then SetWinLoss
                                
                                Exit For
                            Else
                                SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            End If
                        Else
                            If BDoubleNone Then
                                If IPattern(ICounter - 5) = INone Then
                                    SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                Else
                                    SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                End If
                            ElseIf IPattern(ICounter - 5) = INone Then
                                SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            Else
                                If IPattern(ICounter) = ITick Then
                                    If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText2)
                                ElseIf IPattern(ICounter) = ICross Then
                                    If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText1)
                                End If
                                
                                SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                            End If
                        End If
                    End If
                ElseIf IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                    If BDoubleNone Then
                        If IPattern(ICounter - 4) = INone Then
                            If BWinLoss Then SetWinLoss
                            
                            Exit For
                        Else
                            SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        End If
                    Else
                        If IPattern(ICounter - 3) = INone Then
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        ElseIf IPattern(ICounter - 4) = INone Then
                            SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            If BWinLoss Then SetWinLoss
                            
                            Exit For
                        End If
                    End If
                Else
                    If BDoubleNone Then
                        If (IPattern(ICounter - 4) = INone) Or (IPattern(ICounter - 5) = INone) Then
                            SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        End If
                    ElseIf IPattern(ICounter - 4) = INone Then
                        SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                    Else
                        If IPattern(ICounter) = INone Then
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        ElseIf IPattern(ICounter - 3) = INone Then
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        ElseIf IPattern(ICounter) = ITick Then
                            If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText1)
                            
                            SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                        ElseIf IPattern(ICounter) = ICross Then
                            If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText2)
                            
                            SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                        End If
                    End If
                End If
            Else
                If IPattern(ICounter - 2) = IPattern(ICounter - 3) Then
                    If IPattern(ICounter - 2) = INone Then BDoubleNone = True
                
                    If IPattern(ICounter - 3) = IPattern(ICounter - 4) Then
                        If IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                            Exit For
                        ElseIf BDoubleNone Then
                            SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            If BWinLoss Then SetWinLoss
                            
                            Exit For
                        End If
                    ElseIf IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                        If BDoubleNone Then
                            If IPattern(ICounter - 4) = INone Then
                                If BWinLoss Then SetWinLoss
                                
                                Exit For
                            Else
                                SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            End If
                        Else
                            If IPattern(ICounter) = INone Then
                                SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            ElseIf IPattern(ICounter - 1) = INone Then
                                SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            ElseIf IPattern(ICounter - 4) = INone Then
                                SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            ElseIf IPattern(ICounter) = ITick Then
                                If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText2)
                                
                                SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                            ElseIf IPattern(ICounter) = ICross Then
                                If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText1)
                                
                                SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                            End If
                        End If
                    Else
                        If IPattern(ICounter - 1) = INone Then
                            If UBound(IPattern) = 5 Then
                                If BWinLoss Then SetWinLoss
                                'Need Prediction
                                Exit For
                            Else
                                SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            End If
                        ElseIf IPattern(ICounter - 4) = INone Then
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        ElseIf IPattern(ICounter - 5) = INone Then
                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                        Else
                            If BWinLoss Then SetWinLoss
                            
                            Exit For
                        End If
                    End If
                Else
                    If IPattern(ICounter - 3) = IPattern(ICounter - 4) Then
                        If IPattern(ICounter - 3) = INone Then BDoubleNone = True
                        
                        If IPattern(ICounter - 4) = IPattern(ICounter - 5) Then
                            If BDoubleNone Then
                                SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            Else
                                If BWinLoss Then SetWinLoss
                                
                                Exit For
                            End If
                        Else
                            If IPattern(ICounter - 1) = INone Then
                                If BDoubleNone Then
                                    SValue = Set3rdZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                Else
                                    SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                End If
                            Else
                                If BDoubleNone Then
                                    SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                Else
                                    If IPattern(ICounter) = IPattern(ICounter - 5) Then
                                        If IPattern(ICounter) = ITick Then
                                            If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText1)
                                        ElseIf IPattern(ICounter) = ICross Then
                                            If ICounter = UBound(IPattern) Then SValue = SetText(IType, IText2)
                                        End If
                                        
                                        SetPatternResult SValue, IType, BScore, BFocus, BWinLoss, True
                                    Else
                                        If IPattern(ICounter) = INone Then
                                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                        ElseIf IPattern(ICounter - 2) = INone Then
                                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                        ElseIf IPattern(ICounter - 5) = INone Then
                                            SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                        Else
                                            If BWinLoss Then SetWinLoss
                                            
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If IPattern(ICounter) = IPattern(ICounter - 2) Then
                            If IPattern(ICounter) = INone Then
                                SValue = Set2ndZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            Else
                                If IPattern(ICounter - 1) = INone Then
                                    SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                ElseIf IPattern(ICounter - 3) = INone Then
                                    SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                ElseIf IPattern(ICounter - 4) = INone Then
                                    SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                Else
                                    If BWinLoss Then SetWinLoss
                                    
                                    Exit For
                                End If
                            End If
                        Else
                            If IPattern(ICounter - 2) = INone Then
                                SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                            ElseIf IPattern(ICounter - 1) = INone Then
                                If UBound(IPattern) = 5 Then
                                    If BWinLoss Then SetWinLoss
                                    'Need Prediction
                                    Exit For
                                Else
                                    SValue = Set1stZero(IPattern, IType, ICounter, BScore, BFocus, BWinLoss)
                                End If
                            Else
'                                    If IPattern(ICounter) = INone Then
'                                        If ICounter < UBound(IPattern) Then
'                                            ''
'                                        End If
'                                    Else
                                If BWinLoss Then SetWinLoss
                                
                                Exit For
'                                    End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If Not (Trim(SValue) = "") Then SOutput = SValue
    Next ICounter
    
    'If Not (Trim(SOutput) = "") Then SOutput = SOutput & SetScore(IType)
    
    CheckProbPattern = SOutput
End Function

Private Sub SetOutcome(ByVal IType As Integer)
    If IType = IBlackType Then
        IBOutcome = IBOutcome + 1
    ElseIf IType = IEvenType Then
        IEOutcome = IEOutcome + 1
    ElseIf IType = IHighType Then
        IHOutcome = IHOutcome + 1
    End If
End Sub

Private Function CheckPredict1stZero(ByRef IPattern() As Integer, ByVal IType As Integer, ByVal ICounter As Integer) As Boolean
    Dim ITemp() As Integer
    Dim ICounterTemp As Integer
    Dim IPivot As Integer
    Dim ICount As Integer
    
    For ICounterTemp = ICounter To LBound(IPattern) Step -1
        If IPattern(ICounterTemp) = INone Then
            IPivot = ICounterTemp
            
            Exit For
        End If
    Next ICounterTemp
    
    If (IPivot + 5) <= UBound(IPattern) Then
        ICount = (ICounter - IPivot) + ((IPivot + 5) - ICounter)
        ReDim ITemp(ICount) As Integer
    End If
End Function

Private Function Set1stZero(ByRef IPattern() As Integer, ByVal IType As Integer, ByVal ICounter As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False) As String
    Dim ITemp() As Integer
    
    If (ICounter - 5) > 0 Then
        ReDim ITemp(6) As Integer
        
        ITemp(0) = IPattern(ICounter - 6)
        ITemp(1) = IPattern(ICounter - 5)
        ITemp(2) = IPattern(ICounter - 4)
        ITemp(3) = IPattern(ICounter - 3)
        ITemp(4) = IPattern(ICounter - 2)
        ITemp(5) = IPattern(ICounter - 1)
        ITemp(6) = IPattern(ICounter)
    Else
        ReDim ITemp(5) As Integer
        
        ITemp(0) = IPattern(ICounter - 5)
        ITemp(1) = IPattern(ICounter - 4)
        ITemp(2) = IPattern(ICounter - 3)
        ITemp(3) = IPattern(ICounter - 2)
        ITemp(4) = IPattern(ICounter - 1)
        ITemp(5) = IPattern(ICounter)
    End If
    
    Dim SValue As String
    Dim BBet As Boolean
    Dim BCheck As Boolean
    
    If ICounter = UBound(IPattern) Then
        BBet = True
        BCheck = True
    Else
        BBet = False
        BCheck = False
    End If
    
    SValue = Check1stZero(ITemp, IType, BBet, BCheck)
    
    SetPatternResult SValue, IType, BScore, BFocus, BWinLoss
    
    If ICounter = UBound(IPattern) Then Set1stZero = SValue
End Function

Private Function Set2ndZero(ByRef IPattern() As Integer, ByVal IType As Integer, ByVal ICounter As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False) As String
    If (ICounter - 6) > 0 Then
        ReDim ITemp(7) As Integer
        
        ITemp(0) = IPattern(ICounter - 7)
        ITemp(1) = IPattern(ICounter - 6)
        ITemp(2) = IPattern(ICounter - 5)
        ITemp(3) = IPattern(ICounter - 4)
        ITemp(4) = IPattern(ICounter - 3)
        ITemp(5) = IPattern(ICounter - 2)
        ITemp(6) = IPattern(ICounter - 1)
        ITemp(7) = IPattern(ICounter)
    ElseIf (ICounter - 5) > 0 Then
        ReDim ITemp(6) As Integer
        
        ITemp(0) = IPattern(ICounter - 6)
        ITemp(1) = IPattern(ICounter - 5)
        ITemp(2) = IPattern(ICounter - 4)
        ITemp(3) = IPattern(ICounter - 3)
        ITemp(4) = IPattern(ICounter - 2)
        ITemp(5) = IPattern(ICounter - 1)
        ITemp(6) = IPattern(ICounter)
    Else
        ReDim ITemp(5) As Integer
        
        ITemp(0) = IPattern(ICounter - 5)
        ITemp(1) = IPattern(ICounter - 4)
        ITemp(2) = IPattern(ICounter - 3)
        ITemp(3) = IPattern(ICounter - 2)
        ITemp(4) = IPattern(ICounter - 1)
        ITemp(5) = IPattern(ICounter)
    End If
    
    Dim SValue As String
    
    SValue = Check2ndZero(ITemp, IType)
    
    SetPatternResult SValue, IType, BScore, BFocus, BWinLoss
    
    If ICounter = UBound(IPattern) Then Set2ndZero = SValue
End Function

Private Function Set3rdZero(ByRef IPattern() As Integer, ByVal IType As Integer, ByVal ICounter As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False) As String
    ReDim ITemp(5) As Integer
    
    ITemp(0) = IPattern(ICounter - 5)
    ITemp(1) = IPattern(ICounter - 4)
    ITemp(2) = IPattern(ICounter - 3)
    ITemp(3) = IPattern(ICounter - 2)
    ITemp(4) = IPattern(ICounter - 1)
    ITemp(5) = IPattern(ICounter)
    
    Dim SValue As String
    
    SValue = Check3rdZero(ITemp, IType)
    
    SetPatternResult SValue, IType, BScore, BFocus, BWinLoss
    
    If ICounter = UBound(IPattern) Then Set3rdZero = SValue
End Function

Private Sub SetPatternResult(ByVal SValue As String, ByVal IType As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False, Optional ByVal BSkip As Boolean = False)
    If Not (Trim(SValue) = "") Or BSkip Then
        If BScore Then SetNumberScore IType
        If BFocus Then SetNumberFocus IType
        If BWinLoss Then SetWinLoss True
    End If
End Sub

Private Function Check1stZero(ByRef IPattern() As Integer, ByVal IType As Integer, Optional ByVal BBet As Boolean = True, Optional ByVal BCheck As Boolean = True) As String
    Check1stZero = ""
    
    Dim ICounter As Integer
    Dim IPatternCounter As Integer
    Dim IOrder As Integer
    
    IOrder = IBlank
    
    For ICounter = UBound(IPattern) To LBound(IPattern) Step -1
        If IPattern(ICounter) = INone Then
            IOrder = ICounter
            
            Exit For
        End If
    Next ICounter
    
    If Not IOrder = IBlank Then
        Dim IZero(2) As Integer
        Dim ITemp() As Integer
        Dim IStart As Integer
        Dim SValue(2) As String
        
        IZero(0) = INone
        IZero(1) = ICross
        IZero(2) = ITick
        
        For ICounter = LBound(IZero) To UBound(IZero)
            SValue(ICounter) = ""
            
            If IZero(ICounter) = INone Then
                If UBound(IPattern) = 5 Then
                    IStart = 0
                    
                    For IPatternCounter = 0 To UBound(IPattern)
                        If IPatternCounter = IOrder Then
                        Else
                            ReDim Preserve ITemp(IStart) As Integer
                            
                            ITemp(IStart) = IPattern(IPatternCounter)
                            
                            IStart = IStart + 1
                        End If
                    Next IPatternCounter
                    
'                    If Not (Trim(CheckZeroPattern(ITemp, IType)) = "") Then
'                        SetNumberFocus IType
'                    End If
                Else
                    IStart = 0
                    
                    For IPatternCounter = 0 To UBound(IPattern)
                        If IPatternCounter = IOrder Then
                        Else
                            ReDim Preserve ITemp(IStart) As Integer
                            
                            ITemp(IStart) = IPattern(IPatternCounter)
                            
                            IStart = IStart + 1
                        End If
                    Next IPatternCounter
                    
                    SValue(ICounter) = CheckZeroPattern(ITemp, IType)
                End If
            Else
                ReDim ITemp(5) As Integer
                
                If UBound(IPattern) > 5 Then
                    IStart = 1
                Else
                    IStart = 0
                End If
                
                For IPatternCounter = IStart To UBound(IPattern)
                    If IPatternCounter = IOrder Then
                        ITemp(IPatternCounter - IStart) = IZero(ICounter)
                    Else
                        ITemp(IPatternCounter - IStart) = IPattern(IPatternCounter)
                    End If
                Next IPatternCounter
                
                SValue(ICounter) = CheckZeroPattern(ITemp, IType)
            End If
        Next ICounter
    End If
    
    Dim IPivot As Integer
    Dim IOutcome As Integer
    Dim BDouble As Boolean
    
    IOutcome = 0
    BDouble = False
    
    For ICounter = 0 To UBound(SValue)
        If Not (SValue(ICounter) = "") Then IOutcome = IOutcome + 1
        
        For IPatternCounter = 0 To UBound(SValue)
            If ICounter = IPatternCounter Then
            Else
                If SValue(ICounter) = SValue(IPatternCounter) Then
                    If Not (SValue(ICounter) = "") Then
                        IPivot = ICounter
                        
                        BDouble = True
                    End If
                End If
            End If
        Next IPatternCounter
    Next ICounter
    
    Dim IValue As Integer
    
    IValue = 0
    
    If BDouble Then
        If BBet Then SetOutcome IType
        
        Check1stZero = SValue(IPivot)
        
        If BBet Then
            If IType = IBlackType Then
                If SValue(IPivot) = SBlackText Then
                    IValue = IText1
                Else
                    IValue = IText2
                End If
            ElseIf IType = IEvenType Then
                If SValue(IPivot) = SEvenText Then
                    IValue = IText1
                Else
                    IValue = IText2
                End If
            ElseIf IType = IHighType Then
                If SValue(IPivot) = SHighText Then
                    IValue = IText1
                Else
                    IValue = IText2
                End If
            End If
        End If
        
        SetText IType, IValue, BBet, BCheck
    Else
        If IOutcome > 1 Then
            SetNumberFocus IType
            
            Check1stZero = ""
        ElseIf SValue(0) = "" And SValue(1) = "" And SValue(2) = "" Then
            Check1stZero = ""
        ElseIf Not (SValue(0) = "") Then
            Check1stZero = SValue(0)
            
            If BBet Then
                If IType = IBlackType Then
                    If SValue(0) = SBlackText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                ElseIf IType = IEvenType Then
                    If SValue(0) = SEvenText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                ElseIf IType = IHighType Then
                    If SValue(0) = SHighText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                End If
            End If
            
            SetText IType, IValue, BBet, BCheck
        ElseIf Not (SValue(1) = "") Then
            Check1stZero = SValue(1)
            
            If BBet Then
                If IType = IBlackType Then
                    If SValue(1) = SBlackText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                ElseIf IType = IEvenType Then
                    If SValue(1) = SEvenText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                ElseIf IType = IHighType Then
                    If SValue(1) = SHighText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                End If
            End If
            
            SetText IType, IValue, BBet, BCheck
        ElseIf Not (SValue(2) = "") Then
            Check1stZero = SValue(2)
            
            If BBet Then
                If IType = IBlackType Then
                    If SValue(2) = SBlackText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                ElseIf IType = IEvenType Then
                    If SValue(2) = SEvenText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                ElseIf IType = IHighType Then
                    If SValue(2) = SHighText Then
                        IValue = IText1
                    Else
                        IValue = IText2
                    End If
                End If
            End If
            
            SetText IType, IValue, BBet, BCheck
        End If
    End If
End Function

Private Function Check2ndZero(ByRef IPattern() As Integer, ByVal IType As Integer) As String
    Dim ICounterX As Integer
    Dim ICounterY As Integer
    Dim IPatternCounter As Integer
    Dim IOrder1 As Integer
    Dim IOrder2 As Integer
    
    Check2ndZero = ""
    
    IOrder1 = IBlank
    IOrder2 = IBlank
    
    For ICounterX = UBound(IPattern) To LBound(IPattern) Step -1
        If IPattern(ICounterX) = INone Then
            If IOrder1 = IBlank Then
                IOrder1 = ICounterX
            ElseIf IOrder2 = IBlank Then
                IOrder2 = ICounterX
                
                Exit For
            End If
        End If
    Next ICounterX
    
    If Not (IOrder1 = IBlank) And Not (IOrder2 = IBlank) Then
        Dim IZero1(2) As Integer
        Dim IZero2(2) As Integer
        Dim ITemp() As Integer
        Dim IStart As Integer
        Dim SValue As String
        
        IZero1(0) = INone
        IZero1(1) = ICross
        IZero1(2) = ITick
        IZero2(0) = INone
        IZero2(1) = ITick
        IZero2(2) = ICross
        
        For ICounterX = LBound(IZero1) To UBound(IZero1)
            SValue = ""
            
            For ICounterY = LBound(IZero2) To UBound(IZero2)
                If IZero1(ICounterX) = INone Then
                    If IZero2(ICounterY) = INone Then
                        If UBound(IPattern) = 7 Then
                        End If
                    End If
                End If
            Next ICounterY
        Next ICounterX
    End If
End Function

Private Function Check3rdZero(ByRef IPattern() As Integer, ByVal IType As Integer) As String
End Function

Private Function CheckZeroPattern(ByRef IPattern() As Integer, ByVal IType As Integer) As String
    Dim BDoubleTick As Boolean
    Dim BDoubleCross As Boolean
    
    If UBound(IPattern) = 5 Then
        If IPattern(UBound(IPattern)) = IPattern(UBound(IPattern) - 1) Then
            If IPattern(UBound(IPattern)) = ITick Then
                BDoubleTick = True
            ElseIf IPattern(UBound(IPattern)) = ICross Then
                BDoubleCross = True
            End If
            
            If IPattern(UBound(IPattern) - 1) = IPattern(UBound(IPattern) - 2) Then
            ElseIf IPattern(UBound(IPattern) - 2) = IPattern(UBound(IPattern) - 3) Then
                If IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                ElseIf IPattern(UBound(IPattern) - 4) = IPattern(UBound(IPattern) - 5) Then
                Else
                    If BDoubleTick Then
                        CheckZeroPattern = SetText(IType, IText2, False, False)
                    ElseIf BDoubleCross Then
                        CheckZeroPattern = SetText(IType, IText1, False, False)
                    End If
                End If
            ElseIf IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
            ElseIf IPattern(UBound(IPattern) - 4) = IPattern(UBound(IPattern) - 5) Then
                If IPattern(UBound(IPattern)) = ITick Then
                    CheckZeroPattern = SetText(IType, IText2, False, False)
                Else
                    CheckZeroPattern = SetText(IType, IText1, False, False)
                End If
            End If
        Else
            If IPattern(UBound(IPattern) - 1) = IPattern(UBound(IPattern) - 2) Then
                If IPattern(UBound(IPattern) - 2) = IPattern(UBound(IPattern) - 3) Then
                ElseIf IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                    If IPattern(UBound(IPattern) - 4) = IPattern(UBound(IPattern) - 5) Then
                    Else
                        If IPattern(UBound(IPattern)) = ITick Then
                            CheckZeroPattern = SetText(IType, IText2, False, False)
                        ElseIf IPattern(UBound(IPattern)) = ICross Then
                            CheckZeroPattern = SetText(IType, IText1, False, False)
                        End If
                    End If
                ElseIf IPattern(UBound(IPattern) - 4) = IPattern(UBound(IPattern) - 5) Then
                Else
                    If IPattern(UBound(IPattern)) = ITick Then
                        CheckZeroPattern = SetText(IType, IText1, False, False)
                    ElseIf IPattern(UBound(IPattern)) = ICross Then
                        CheckZeroPattern = SetText(IType, IText2, False, False)
                    End If
                End If
            Else
                If IPattern(UBound(IPattern) - 2) = IPattern(UBound(IPattern) - 3) Then
                    If IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                    ElseIf IPattern(UBound(IPattern) - 4) = IPattern(UBound(IPattern) - 5) Then
                        If IPattern(UBound(IPattern)) = ITick Then
                            CheckZeroPattern = SetText(IType, IText2, False, False)
                        ElseIf IPattern(UBound(IPattern)) = ICross Then
                            CheckZeroPattern = SetText(IType, IText1, False, False)
                        End If
                    End If
                Else
                    If IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                        If IPattern(UBound(IPattern)) = IPattern(UBound(IPattern) - 5) Then
                            If IPattern(UBound(IPattern)) = ITick Then
                                CheckZeroPattern = SetText(IType, IText1, False, False)
                            ElseIf IPattern(UBound(IPattern)) = ICross Then
                                CheckZeroPattern = SetText(IType, IText2, False, False)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    ElseIf UBound(IPattern) = 4 Then
        If IPattern(UBound(IPattern)) = IPattern(UBound(IPattern) - 1) Then
            If IPattern(UBound(IPattern)) = ITick Then
                BDoubleTick = True
            ElseIf IPattern(UBound(IPattern)) = ICross Then
                BDoubleCross = True
            End If
        
            If IPattern(UBound(IPattern) - 1) = IPattern(UBound(IPattern) - 2) Then
            ElseIf IPattern(UBound(IPattern) - 2) = IPattern(UBound(IPattern) - 3) Then
                If IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                Else
                    If BDoubleTick Then
                        CheckZeroPattern = SetText(IType, IText2, False, False)
                    ElseIf BDoubleCross Then
                        CheckZeroPattern = SetText(IType, IText1, False, False)
                    End If
                End If
            ElseIf IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
            End If
        Else
            If IPattern(UBound(IPattern) - 1) = IPattern(UBound(IPattern) - 2) Then
                If IPattern(UBound(IPattern) - 2) = IPattern(UBound(IPattern) - 3) Then
                ElseIf IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                    If IPattern(UBound(IPattern)) = ITick Then
                        CheckZeroPattern = SetText(IType, IText2, False, False)
                    ElseIf IPattern(UBound(IPattern)) = ICross Then
                        CheckZeroPattern = SetText(IType, IText1, False, False)
                    End If
                End If
            Else
                If IPattern(UBound(IPattern) - 2) = IPattern(UBound(IPattern) - 3) Then
                    If IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                    Else
                        If IPattern(UBound(IPattern)) = ITick Then
                            CheckZeroPattern = SetText(IType, IText2, False, False)
                        ElseIf IPattern(UBound(IPattern)) = ICross Then
                            CheckZeroPattern = SetText(IType, IText1, False, False)
                        End If
                    End If
                Else
                    If IPattern(UBound(IPattern) - 3) = IPattern(UBound(IPattern) - 4) Then
                        If IPattern(UBound(IPattern)) = ITick Then
                            CheckZeroPattern = SetText(IType, IText1, False, False)
                        ElseIf IPattern(UBound(IPattern)) = ICross Then
                            CheckZeroPattern = SetText(IType, IText2, False, False)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub SetNumberScore(ByVal IType As Integer)
    If IType = IBlackType Then
        IBlackScore = IBlackScore + 1
    ElseIf IType = IEvenType Then
        IEvenScore = IEvenScore + 1
    ElseIf IType = IHighType Then
        IHighScore = IHighScore + 1
    End If
End Sub

Private Sub SetNumberFocus(ByVal IType As Integer)
    If IType = IBlackType Then
        IBlackFocus = IBlackFocus + 1
    ElseIf IType = IEvenType Then
        IEvenFocus = IEvenFocus + 1
    ElseIf IType = IHighType Then
        IHighFocus = IHighFocus + 1
    End If
End Sub

Private Function SetText(ByVal IType As Integer, ByVal IValue As Integer, Optional ByVal BBet As Boolean = True, Optional ByVal BCheck As Boolean = True)
    SetText = ""
    
    If IType = IBlackType Then
        If IValue = IText1 Then
            SetText = SBlackText
            
            If BBet Then BBlackPattern = True
        ElseIf IValue = IText2 Then
            SetText = SRedText
            
            If BBet Then BBlackPattern = False
        End If
        
        If BCheck Then BBlack = True
    ElseIf IType = IEvenType Then
        If IValue = IText1 Then
            SetText = SEvenText
            
            If BBet Then BEvenPattern = True
        ElseIf IValue = IText2 Then
            SetText = SOddText
            
            If BBet Then BEvenPattern = False
        End If
        
        If BCheck Then BEven = True
    ElseIf IType = IHighType Then
        If IValue = IText1 Then
            SetText = SHighText
            
            If BBet Then BHighPattern = True
        ElseIf IValue = IText2 Then
            SetText = SLowText
            
            If BBet Then BHighPattern = False
        End If
        
        If BCheck Then BHigh = True
    End If
End Function

Private Function SetScore(ByVal IType As Integer) As String
    SetScore = ""
    
    If IType = IBlackType Then
        If Not IBlackScore = 0 Then
            SetScore = " (" & IBlackScore & ")"
        End If
    ElseIf IType = IEvenType Then
        If Not IEvenScore = 0 Then
            SetScore = " (" & IEvenScore & ")"
        End If
    ElseIf IType = IHighType Then
        If Not IHighScore = 0 Then
            SetScore = " (" & IHighScore & ")"
        End If
    End If
End Function

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
