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
    
    SBlack1 = CheckProbPattern(IBlack1RollPattern, IBlack1Type)
    
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

Public Function CheckProbPattern(ByRef IPattern() As Integer, ByVal IType As Integer) As String
    Dim ICounter As Integer
    Dim BTripleBlackTick As Boolean
    Dim BTripleBlackCross As Boolean
    Dim SOutput As String
    Dim SValue As String
    
    CheckProbPattern = ""
    SOutput = ""
        
    Dim BValid As Boolean
    
    BValid = False
    
    If UBound(IPattern) < 2 Then
        Exit Function
    End If
    
    SValue = ""
    BTripleBlackTick = False
    BTripleBlackCross = False
    
    ICounter = UBound(IPattern)
    
    If IPattern(ICounter) = IPattern(ICounter - 1) Then
        If IPattern(ICounter) = IPattern(ICounter - 2) Then
            If IPattern(ICounter) = ITick Then
                BTripleBlackTick = True
            ElseIf IPattern(ICounter) = ICross Then
                BTripleBlackCross = True
            End If
        End If
    End If
    
    If BTripleBlackTick Then
        If IType = IBlack1Type Then
            If IBlack1Focus = -1 Then
                IBlack1Focus = UBound(IPattern) - 2
            End If
            
            BBlack1 = True
        End If
        
        SValue = SetText(IType, IText1)
        BValid = True
    ElseIf BTripleBlackCross Then
        If IType = IBlack1Type Then
            If IBlack1Focus = -1 Then
                IBlack1Focus = UBound(IPattern) - 2
            End If
            
            BBlack1 = True
        End If
    
        SValue = SetText(IType, IText2)
        BValid = True
    End If
    
    If BValid Then
        CheckProbPattern = SValue
        Exit Function
    End If
    
    Dim BZigZag As Boolean
    SValue = ""
    BZigZag = False
    
    ICounter = UBound(IPattern)
    
    If IPattern(ICounter) = ITick And IPattern(ICounter - 1) = ICross And IPattern(ICounter - 2) = ITick Then
        BZigZag = True
    ElseIf IPattern(ICounter) = ICross And IPattern(ICounter - 1) = ITick And IPattern(ICounter - 2) = ICross Then
        BZigZag = True
    End If
    
    If BZigZag Then
        If IType = IBlack1Type Then
            If IBlack1Focus = -1 Then
                IBlack1Focus = UBound(IPattern) - 2
            End If
            
            BBlack1 = True
        End If
        
        SValue = SetText(IType, IText2)
        BValid = True
    End If
    
    If BValid Then
        CheckProbPattern = SValue
        Exit Function
    End If
    
    If UBound(IPattern) > 4 Then
        Dim BDoubleTag As Boolean
        SValue = ""
        BDoubleTag = False
        
        ICounter = UBound(IPattern)
        
        If IPattern(ICounter) = ICross And IPattern(ICounter - 1) = ICross And IPattern(ICounter - 2) = ITick And IPattern(ICounter - 3) = ITick Then
            BDoubleTag = True
        End If
        
        If BDoubleTag Then
            If IType = IBlack1Type Then
                If IBlack1Focus = -1 Then
                    IBlack1Focus = UBound(IPattern) - 3
                End If
                
                BBlack1 = True
            End If
            
            SValue = SetText(IType, IText1)
            BValid = True
        End If
    End If
    
    If Trim(SValue) = "" Then
        If IType = IBlack1Type Then
            IBlack1Focus = -1
            
            BBlack1 = False
        End If
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
    End If
End Function
