Attribute VB_Name = "mdSecurity"
Option Explicit

Public Const SKey As String = "ROULETTE"
Public Const SDSecure As String = "ú>ŽZóô3Ò’@Qªý#" '149758018
Public Const SDGuest As String = "ä7In0á"

Public SSecured As String
Public SGuest As String

Private Const SSecureText As String = "SECURE"
Private Const SGuestStartText As String = "GUEST"

Private bytKey(21, 8) As Byte
Private bytKey1(8) As Byte
Private bytMessage() As Byte
Private bytResult() As Byte
Private exptab(255) As Byte
Private logtab(255) As Byte

Private blnGuest As Boolean

Private Const intIteration As Integer = 6

Public Sub Init()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO)
    
    Dim SValue As String
    
    lRegistry = mdAPI.ReadValueRegistry(lRegKey, SSecureText, lType, SValue, lSize)
    
    SValue = mdAPI.ReplaceRegistry(SValue)
    
    If Trim(SValue) = "" Then
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, SSecureText, SDSecure)
        
        SSecured = SDSecure
    Else
        SSecured = SValue
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub InitGuest()
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO)
    
    Dim SValue As String
    
    lRegistry = mdAPI.ReadValueRegistry(lRegKey, SGuestStartText, lType, SValue, lSize)
    
    SValue = mdAPI.ReplaceRegistry(SValue)
    
    If Trim(SValue) = "" Then
        lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, SGuestStartText, Now)
        
        SGuest = DateAdd("h", 100, Now)
    Else
        SGuest = DateAdd("h", 100, SValue)
    End If
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Sub ChangePassword(ByVal strPassword As String)
    On Local Error GoTo ErrHandler
    
    Dim lRegistry As Long
    Dim lRegKey As Long
    Dim lType As Long
    Dim lSize As Long
    
    lRegistry = mdAPI.OpenRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, lRegKey)
    
    If Not lRegistry = 0 Then lRegistry = mdAPI.WriteToRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO)
    
    lRegistry = mdAPI.WriteValueRegistry(mdAPI.HKEY_CURRENT_USER, mdAPI.KEYS_SYS_INFO, SSecureText, EncryptText(strPassword, SKey))
    
    lRegistry = mdAPI.CloseRegistry(lRegKey)
    
    Exit Sub
ErrHandler:
End Sub

Public Function IsGuest(Optional ByVal blnValue As Boolean = True) As Boolean
    If blnValue Then
        InitGuest
        
        If DateDiff("s", Now, SGuest) >= 30 Then
            IsGuest = True
        Else
            IsGuest = False
        End If
    Else
        IsGuest = False
    End If

    blnGuest = blnValue
End Function

Public Function GetGuest() As Boolean
    GetGuest = blnGuest
End Function

Public Function EncryptText(ByVal strMessage As String, ByVal strKey As String) As String
    If Not Len(strMessage) Mod 8 = 0 Then
        strMessage = strMessage & Space(8 - (Len(strMessage) Mod 8))
    End If
    
    If Len(strKey) > 8 Then
        EncryptText = ""
        
        Exit Function
    End If
    
    If Not Len(strKey) Mod 8 = 0 Then
        strKey = strKey & Space(8 - (Len(strKey) Mod 8))
    End If
    
    SaferInit
    KeyUserInit strKey
    KeySchedule
    EncryptText = SaferK64Encrypt(strMessage)
    'Debug.Print EncryptText
End Function

Private Function SaferK64Encrypt(ByVal strMessage As String) As String
    Dim intCounter As Integer
    
    Dim bytTemp() As Byte
    
    Dim strTemp As String
    Dim strResult As String
    
    ReDim bytResult(7, 0) As Byte
    
    strResult = ""
    
    For intCounter = 1 To Len(strMessage) Step 8
        strTemp = _
            Mid(strMessage, intCounter, 1) & _
            Mid(strMessage, intCounter + 1, 1) & _
            Mid(strMessage, intCounter + 2, 1) & _
            Mid(strMessage, intCounter + 3, 1) & _
            Mid(strMessage, intCounter + 4, 1) & _
            Mid(strMessage, intCounter + 5, 1) & _
            Mid(strMessage, intCounter + 6, 1) & _
            Mid(strMessage, intCounter + 7, 1)
            
        bytMessage = StringToByte(strTemp)

        bytTemp = SaferEncrypt(intCounter \ 8)

        strResult = strResult & ByteToString(bytTemp)
    Next intCounter
    
    SaferK64Encrypt = strResult
End Function

Private Function SaferEncrypt(Optional ByVal intCounter As Integer = 0) As Byte()
    Dim b1 As Byte
    Dim b2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte
    Dim b5 As Byte
    Dim b6 As Byte
    Dim b7 As Byte
    Dim b8 As Byte
    
    Dim intI As Integer

    For intI = 1 To intIteration
        bytMessage(0) = Conversion.CByte((bytMessage(0) Xor bytKey(2 * intI - 1, 1)) Mod 256)
        bytMessage(1) = Conversion.CByte((Conversion.CInt(bytMessage(1)) + Conversion.CInt(bytKey(2 * intI - 1, 2))) Mod 256)
        bytMessage(2) = Conversion.CByte((Conversion.CInt(bytMessage(2)) + Conversion.CInt(bytKey(2 * intI - 1, 3))) Mod 256)
        bytMessage(3) = Conversion.CByte((bytMessage(3) Xor bytKey(2 * intI - 1, 4)) Mod 256)
        bytMessage(4) = Conversion.CByte((bytMessage(4) Xor bytKey(2 * intI - 1, 5)) Mod 256)
        bytMessage(5) = Conversion.CByte((Conversion.CInt(bytMessage(5)) + Conversion.CInt(bytKey(2 * intI - 1, 6))) Mod 256)
        bytMessage(6) = Conversion.CByte((Conversion.CInt(bytMessage(6)) + Conversion.CInt(bytKey(2 * intI - 1, 7))) Mod 256)
        bytMessage(7) = Conversion.CByte((bytMessage(7) Xor bytKey(2 * intI - 1, 8)) Mod 256)

        b1 = Conversion.CByte(exptab(bytMessage(0)) Mod 256)
        b2 = Conversion.CByte(logtab(bytMessage(1)) Mod 256)
        b3 = Conversion.CByte(logtab(bytMessage(2)) Mod 256)
        b4 = Conversion.CByte(exptab(bytMessage(3)) Mod 256)
        b5 = Conversion.CByte(exptab(bytMessage(4)) Mod 256)
        b6 = Conversion.CByte(logtab(bytMessage(5)) Mod 256)
        b7 = Conversion.CByte(logtab(bytMessage(6)) Mod 256)
        b8 = Conversion.CByte(exptab(bytMessage(7)) Mod 256)

        b1 = Conversion.CByte((Conversion.CInt(b1) + Conversion.CInt(bytKey(2 * intI, 1))) Mod 256)
        b2 = Conversion.CByte((b2 Xor bytKey(2 * intI, 2)) Mod 256)
        b3 = Conversion.CByte((b3 Xor bytKey(2 * intI, 3)) Mod 256)
        b4 = Conversion.CByte((Conversion.CInt(b4) + Conversion.CInt(bytKey(2 * intI, 4))) Mod 256)
        b5 = Conversion.CByte((Conversion.CInt(b5) + Conversion.CInt(bytKey(2 * intI, 5))) Mod 256)
        b6 = Conversion.CByte((b6 Xor bytKey(2 * intI, 6)) Mod 256)
        b7 = Conversion.CByte((b7 Xor bytKey(2 * intI, 7)) Mod 256)
        b8 = Conversion.CByte((Conversion.CInt(b8) + Conversion.CInt(bytKey(2 * intI, 8))) Mod 256)

        Mat1 b1, b2, bytMessage(0), bytMessage(1)
        Mat1 b3, b4, bytMessage(2), bytMessage(3)
        Mat1 b5, b6, bytMessage(4), bytMessage(5)
        Mat1 b7, b8, bytMessage(6), bytMessage(7)
        
        Mat1 bytMessage(0), bytMessage(2), b1, b2
        Mat1 bytMessage(4), bytMessage(6), b3, b4
        Mat1 bytMessage(1), bytMessage(3), b5, b6
        Mat1 bytMessage(5), bytMessage(7), b7, b8

        Mat1 b1, b3, bytMessage(0), bytMessage(1)
        Mat1 b5, b7, bytMessage(2), bytMessage(3)
        Mat1 b2, b4, bytMessage(4), bytMessage(5)
        Mat1 b6, b8, bytMessage(6), bytMessage(7)

        ReDim Preserve bytResult(7, intCounter) As Byte

        bytResult(0, intCounter) = bytMessage(0)
        bytResult(1, intCounter) = bytMessage(1)
        bytResult(2, intCounter) = bytMessage(2)
        bytResult(3, intCounter) = bytMessage(3)
        bytResult(4, intCounter) = bytMessage(4)
        bytResult(5, intCounter) = bytMessage(5)
        bytResult(6, intCounter) = bytMessage(6)
        bytResult(7, intCounter) = bytMessage(7)
    Next intI

    bytMessage(0) = Conversion.CByte((bytMessage(0) Xor bytKey(2 * intIteration + 1, 1)) Mod 256)
    bytMessage(1) = Conversion.CByte((Conversion.CInt(bytMessage(1)) + Conversion.CInt(bytKey(2 * intIteration + 1, 2))) Mod 256)
    bytMessage(2) = Conversion.CByte((Conversion.CInt(bytMessage(2)) + Conversion.CInt(bytKey(2 * intIteration + 1, 3))) Mod 256)
    bytMessage(3) = Conversion.CByte((bytMessage(3) Xor bytKey(2 * intIteration + 1, 4)) Mod 256)
    bytMessage(4) = Conversion.CByte((bytMessage(4) Xor bytKey(2 * intIteration + 1, 5)) Mod 256)
    bytMessage(5) = Conversion.CByte((Conversion.CInt(bytMessage(5)) + Conversion.CInt(bytKey(2 * intIteration + 1, 6))) Mod 256)
    bytMessage(6) = Conversion.CByte((Conversion.CInt(bytMessage(6)) + Conversion.CInt(bytKey(2 * intIteration + 1, 7))) Mod 256)
    bytMessage(7) = Conversion.CByte((bytMessage(7) Xor bytKey(2 * intIteration + 1, 8)) Mod 256)

    SaferEncrypt = bytMessage
End Function

Private Sub Mat1( _
    ByRef a1 As Byte, _
    ByRef a2 As Byte, _
    ByRef b1 As Byte, _
    ByRef b2 As Byte)
    b2 = Conversion.CByte((Conversion.CInt(a1) + Conversion.CInt(a2)) Mod 256)
    b1 = Conversion.CByte((Conversion.CInt(b2) + Conversion.CInt(a1)) Mod 256)
End Sub

Private Sub InvMat1( _
    ByRef a1 As Byte, _
    ByRef a2 As Byte, _
    ByRef b1 As Byte, _
    ByRef b2 As Byte)
    b1 = SetLimited((Conversion.CInt(a1) - Conversion.CInt(a2)) Mod 256)
    b2 = SetLimited((Conversion.CInt(-(b1)) + Conversion.CInt(a2)) Mod 256)
End Sub

Private Function SetLimited(ByVal intTemp As Integer) As Byte
    If intTemp < 0 Then
        SetLimited = Conversion.CByte(intTemp + 256)
    Else
        SetLimited = Conversion.CByte(intTemp)
    End If
End Function

Private Sub SaferInit()
    Dim intTemp As Integer
    
    intTemp = 1
    
    logtab(1) = 0
    exptab(0) = 1
    
    Dim intCounter As Integer
    
    For intCounter = 1 To 255
        intTemp = (45 * intTemp) Mod 257
        
        If Not intCounter = 128 Then
            exptab(intCounter) = Conversion.CByte(intTemp)
            
            logtab(exptab(intCounter)) = Conversion.CByte(intCounter)
        End If
    Next intCounter
    
    exptab(128) = 0
    logtab(0) = 128
End Sub

Private Sub KeyUserInit(ByVal strKey As String)
    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strKey)
        bytKey1(intCounter) = Conversion.CByte(Asc(Mid(strKey, intCounter, 1)))
        
        bytKey(1, intCounter) = bytKey1(intCounter)
    Next intCounter
End Sub

Private Sub KeySchedule()
    Dim intTemp As Long
    
    Dim intI As Integer
    Dim intJ As Integer
    
    For intI = 2 To (2 * intIteration + 1)
        For intJ = 1 To 8
            bytKey1(intJ) = Conversion.CByte((Shl(bytKey1(intJ), 3) + Shr(bytKey1(intJ), 5)) And 255)
            
            intTemp = Conversion.CInt(bytKey1(intJ)) + Conversion.CInt(exptab(exptab(9 * intI + intJ)))
            
            bytKey(intI, intJ) = Conversion.CByte(intTemp Mod 256)
        Next intJ
    Next intI
End Sub

Private Function Shl(ByVal bytValue As Byte, ByVal intValue As Integer) As Integer
    Dim strBinary As String
    Dim strTemp As String

    strBinary = ByteToBinary(bytValue)
    strTemp = ""
    
    While Not Len(strTemp) = intValue
        strTemp = strTemp & "0"
    Wend
    
    strBinary = Mid(strBinary, intValue) & strTemp
    
    Shl = BinaryToInteger(strBinary)
End Function

Private Function Shr(ByVal bytValue As Byte, ByVal intValue As Integer) As Integer
    Dim strBinary As String
    Dim strTemp As String
    
    strBinary = ByteToBinary(bytValue)
    strTemp = ""
    
    While Not Len(strTemp) = intValue
        strTemp = strTemp & "0"
    Wend
    
    strBinary = strTemp & Mid(strBinary, 1, Len(strBinary) - Len(strTemp) - 1)
    
    Shr = BinaryToInteger(strBinary)
End Function

Private Function StringToByte(ByVal strValue As String) As Byte()
    Dim bytResult() As Byte
    
    ReDim bytResult(Len(strValue) - 1) As Byte
    
    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strValue)
        bytResult(intCounter - 1) = Conversion.CByte(Asc(Mid(strValue, intCounter, 1)))
    Next intCounter

    StringToByte = bytResult
End Function

Private Function ByteToString(ByRef bytValue() As Byte) As String
    Dim strResult As String
    
    strResult = ""
    
    Dim intCounter As Integer
    
    For intCounter = 0 To UBound(bytValue)
        strResult = strResult & Chr(bytValue(intCounter))
    Next intCounter
    
    ByteToString = strResult
End Function

Private Function ByteToBinary(ByVal bytValue As Byte) As String
    Dim intTemp As Integer

    Dim strResult As String
    
    strResult = ""
    
    Do
        intTemp = bytValue Mod 2
        
        strResult = Conversion.CStr(intTemp) & strResult
        
        bytValue = Conversion.CByte(bytValue \ 2)
    Loop Until bytValue = 0
    
    While Len(strResult) < 32
        strResult = "0" & strResult
    Wend
    
    ByteToBinary = strResult
End Function

Private Function BinaryToInteger(ByVal strBinary As String) As Integer
    Dim intTarget As Integer
    
    intTarget = 0
    
    Dim intCounter As Integer
    
    For intCounter = Len(strBinary) To 1 Step -1
        intTarget = _
            intTarget + _
            EachBinaryToInteger(Mid(strBinary, intCounter, 1), Len(strBinary) - intCounter)
    Next intCounter
    
    BinaryToInteger = intTarget
End Function

Private Function EachBinaryToInteger(ByVal strEachBinary As String, ByVal intCounter As Integer) As Integer
    Select Case strEachBinary
        Case "0"
            EachBinaryToInteger = 0
        Case "1"
            EachBinaryToInteger = CInt(2 ^ intCounter)
    End Select
End Function


