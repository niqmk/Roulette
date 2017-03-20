Attribute VB_Name = "mdPattern"
Option Explicit

Private Const IText1 As Integer = 1
Private Const IText2 As Integer = 2

Private BBlack1Pattern As Boolean

Public Function CheckPattern(Optional ByVal BAlert As Boolean = True) As String
    Dim SBlack1 As String
    Dim SBlack2 As String
    Dim SBlack3 As String
    Dim SBlack4 As String
    
    SBlack1 = CheckProbPattern(mdApp.IBlack1RollPattern)
    
    Dim SValue As String
    
    SValue = SBlack1
    SValue = SValue & vbCrLf
    SValue = SValue & SBlack2
    SValue = SValue & vbCrLf
    SValue = SValue & SBlack3
    SValue = SValue & vbCrLf
    SValue = SValue & SBlack4
    
    If BAlert Then
        If Trim(SBlack1) = "" And Trim(SBlack2) = "" And Trim(SBlack3) = "" And Trim(SBlack4) = "" Then
            mdAPI.Beep 300, 50
        Else
            mdAPI.Beep 700, 300
        End If
    End If

    CheckPattern = SValue
End Function

Public Function CheckProbPattern(ByRef IPattern() As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False) As String
    Dim ICounter As Integer
    Dim BTripleBlackTick As Boolean
    Dim BTripleBlackCross As Boolean
    Dim SOutput As String
    Dim SValue As String
    
    CheckProbPattern = ""
    
    If UBound(IPattern) < 2 Then
        Exit Function
    End If
    
    SOutput = ""
    
    Dim BValid As Boolean
    
    BValid = False
    
    For ICounter = UBound(IPattern) To UBound(IPattern) - 3 Step -4
        SValue = ""
        BTripleBlackTick = False
        BTripleBlackCross = False
        
        If IPattern(ICounter) = IPattern(ICounter - 1) Then
            If IPattern(ICounter) = IPattern(ICounter - 2) Then
                If IPattern(ICounter) = mdApp.ITick Then
                    BTripleBlackTick = True
                ElseIf IPattern(ICounter) = mdApp.ICross Then
                    BTripleBlackCross = True
                End If
            End If
        End If
    Next ICounter
    
    If BTripleBlackTick Then
        SValue = SetText(IBlackType, IText1)
        BValid = True
    ElseIf BTripleBlackCross Then
        SValue = SetText(IBlackType, IText2)
        BValid = True
    End If
    
    If BValid Then
        CheckProbPattern = SValue
        Exit Function
    End If
    
    Dim BZigZag As Boolean
    
    For ICounter = UBound(IPattern) To UBound(IPattern) - 3 Step -4
        SValue = ""
        BZigZag = False
        
        If IPattern(ICounter) = mdApp.ITick And IPattern(ICounter - 1) = mdApp.ICross And IPattern(ICounter - 2) = mdApp.ITick Then
            BZigZag = True
        End If
    Next ICounter
    
    If BZigZag Then
        SValue = SetText(IBlackType, IText2)
        BValid = True
    End If
    
    If BValid Then
        CheckProbPattern = SValue
        Exit Function
    End If
    
    Dim BDoubleTag As Boolean
    
    For ICounter = UBound(IPattern) To UBound(IPattern) - 4 Step -5
        SValue = ""
        BDoubleTag = False
        
        If IPattern(ICounter) = mdApp.ICross And IPattern(ICounter - 1) = mdApp.ICross And IPattern(ICounter - 2) = mdApp.ITick And IPattern(ICounter - 3) = mdApp.ITick Then
            BDoubleTag = True
        End If
    Next ICounter
    
    If BDoubleTag Then
        SValue = SetText(IBlackType, IText1)
        BValid = True
    End If
    
    CheckProbPattern = SValue
End Function

Private Function SetText(ByVal IType As Integer, ByVal IValue As Integer, Optional ByVal BBet As Boolean = True, Optional ByVal BCheck As Boolean = True)
    SetText = ""
    
    If IType = IBlackType Then
        If IValue = IText1 Then
            SetText = SBlackText
            
            If BBet Then BBlack1Pattern = True
        ElseIf IValue = IText2 Then
            SetText = SRedText
            
            If BBet Then BBlack1Pattern = False
        End If
        
        If BCheck Then mdApp.BBlack1 = True
    End If
End Function
