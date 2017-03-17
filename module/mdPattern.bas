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

Private Function CheckProbPattern(ByRef IPattern() As Integer, Optional ByVal BScore As Boolean = True, Optional ByVal BFocus As Boolean = True, Optional ByVal BWinLoss As Boolean = False) As String
    Dim ICounter As Integer
    Dim BTripleBlackTick As Boolean
    Dim BTripleBlackCross As Boolean
    Dim SOutput As String
    Dim SValue As String
    
    CheckProbPattern = ""
    
    If UBound(IPattern) < 3 Then
        Exit Function
    End If
    
    SOutput = ""
    
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
        SValue = SetText(IBlack1Type, IText1)
    ElseIf BTripleBlackCross Then
        SValue = SetText(IBlack1Type, IText2)
    End If
    
    CheckProbPattern = SValue
End Function

Private Function SetText(ByVal IType As Integer, ByVal IValue As Integer, Optional ByVal BBet As Boolean = True, Optional ByVal BCheck As Boolean = True)
    SetText = ""
    
    If IType = IBlack1Type Then
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
