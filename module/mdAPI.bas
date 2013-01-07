Attribute VB_Name = "mdAPI"
Option Explicit

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDesc As Long
    bInheritHandle As Long
End Type

Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD As Long = &H0
Public Const NIM_MODIFY As Long = &H1
Public Const NIM_DELETE As Long = &H2
Public Const NIF_MESSAGE As Long = &H1
Public Const NIF_ICON As Long = &H2
Public Const NIF_TIP As Long = &H4
Public Const WM_MOUSEMOVE As Long = &H200
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
Public Const WM_LBUTTONDBLCLK As Long = &H203
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_RBUTTONDBLCLK As Long = &H206

Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1

Private SECURITY_ATT As SECURITY_ATTRIBUTES

Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_CREATE_LINK As Long = &H20
Public Const KEY_READ As Long = &H20019
Public Const KEY_WRITE As Long = &H20006
Public Const KEY_ALL_ACCESS As Long = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + &H20000

Public Const REG_SZ As Integer = 1
Public Const REG_DWORD As Integer = 4

Public Const KEYS_SYS_INFO As String = "SOFTWARE\MJSTONE\RS QUICK\"
Public Const KEYS_SYS_HISTORY As String = "SOFTWARE\MJSTONE\RS QUICK\HISTORY\"
Public Const KEYS_SYS_TABLE As String = "SOFTWARE\MJSTONE\RS QUICK\TABLE\"
Public Const KEYS_SYS_BOX As String = "SOFTWARE\MJSTONE\RS QUICK\BOX\"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal lKey As Long, ByVal strSubKey As String, ByVal lOptions As Long, ByVal lDesired As Long, ByRef lResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal lKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As Any, ByRef lpcbData As Long) As Long

Private Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Function WriteToRegistry(ByRef lKeyRoot As Long, ByRef SKeyName As String) As Long
    Dim lRegKey As Long

    SECURITY_ATT.lpSecurityDesc = 0
    SECURITY_ATT.bInheritHandle = True
    SECURITY_ATT.nLength = Len(SECURITY_ATT)

    WriteToRegistry = RegCreateKeyEx(lKeyRoot, SKeyName, 0, "", 0, KEY_ALL_ACCESS, SECURITY_ATT, lRegKey, 0)
End Function

Public Function WriteValueRegistry(ByRef lKeyRoot As Long, ByRef SKeyName As String, ByRef SSubKey As String, ByRef SValue As String) As Long
    Dim lRegKey As Long

    WriteValueRegistry = RegCreateKey(lKeyRoot, SKeyName, lRegKey)
    WriteValueRegistry = RegSetValueEx(lRegKey, SSubKey, 0, REG_SZ, ByVal SValue, Len(SValue))
End Function

Public Function OpenRegistry(ByRef lKeyRoot As Long, ByRef SKeyName As String, ByRef lValue As Long) As Long
    Dim lRegKey As Long

    OpenRegistry = RegOpenKeyEx(lKeyRoot, SKeyName, 0, KEY_ALL_ACCESS, lRegKey)
    
    lValue = lRegKey
End Function

Public Function ReadValueRegistry(ByRef lKeyValueRegistry As Long, ByRef SSubKey As String, ByRef lTypeBack As Long, ByRef SDataBack As String, ByRef lSizeBack As Long) As Long
    SDataBack = String$(1024, 0)
    lSizeBack = 1024

    ReadValueRegistry = RegQueryValueEx(lKeyValueRegistry, SSubKey, 0, lTypeBack, SDataBack, lSizeBack)
End Function

Public Function DeleteKeysRegistry(ByRef lKeyRoot As Long, ByRef SKeyName As String) As Long
    DeleteKeysRegistry = RegDeleteKey(lKeyRoot, SKeyName)
End Function

Public Function DeleteSubKeysRegistry(ByRef lKeyRoot As Long, ByRef SKeyName As String, ByRef SSubKey As String) As Long
    Dim lRegKey As Long

    DeleteSubKeysRegistry = RegOpenKeyEx(lKeyRoot, SKeyName, 0, KEY_ALL_ACCESS, lRegKey)
    
    If DeleteSubKeysRegistry = 0 Then DeleteSubKeysRegistry = RegDeleteValue(lRegKey, SSubKey)
End Function

Public Function CloseRegistry(ByRef lKeyValueRegistry As Long) As Long
    CloseRegistry = RegCloseKey(lKeyValueRegistry)
End Function

Public Function ReplaceRegistry(ByVal SText As String) As String
    Dim IPos As String
    
    IPos = InStr(SText, Chr(0))
    
    If Not IPos = 0 Then
        SText = Mid(SText, 1, IPos - 1)
    Else
        SText = ""
    End If
    
    ReplaceRegistry = SText
End Function

Public Sub ShapeForm(ByRef oForm As Form, ByVal SText As String, ByVal LColor As Long, ByVal lHeight As Long, Optional ByVal sX As Single = 0, Optional ByVal sY As Single = 0)
    Const FW_BOLD As Long = 300
    Dim lNFont As Long
    Dim lOFont As Long
    Dim lhRgn As Long

    oForm.AutoRedraw = True
    oForm.BorderStyle = vbBSNone
    oForm.ScaleMode = vbPixels
    oForm.BackColor = LColor

    lNFont = CustomFont(lHeight, 0, 0, 0, 700, False, False, False, "Times New Roman")
    lOFont = SelectObject(oForm.hDC, lNFont)

    SelectObject oForm.hDC, lNFont
    BeginPath oForm.hDC
    
    oForm.CurrentX = sX
    oForm.CurrentY = sY
    oForm.Print SText
    EndPath oForm.hDC
    
    lhRgn = PathToRegion(oForm.hDC)

    SetWindowRgn oForm.hWnd, lhRgn, False

    SelectObject oForm.hDC, lOFont

    DeleteObject lNFont
End Sub

Private Function CustomFont(ByVal lHeight As Long, ByVal lWidth As Long, ByVal lEscapement As Long, ByVal lOrientation As Long, ByVal lWeight As Long, ByVal lIsItalic As Long, ByVal lIsUnderscored As Long, ByVal lIsStrikenOut As Long, ByVal SFace As String) As Long
    Const CLIP_LH_ANGLES = 16
    
    CustomFont = CreateFont(lHeight, lWidth, lEscapement, lOrientation, lWeight, lIsItalic, lIsUnderscored, lIsStrikenOut, 0, 0, CLIP_LH_ANGLES, 0, 0, SFace)
End Function
