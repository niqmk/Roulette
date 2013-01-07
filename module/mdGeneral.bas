Attribute VB_Name = "mdGeneral"
Option Explicit

Public SPath As String

Public Sub CenterWindows(ByRef frmCenter As Form, Optional ByVal blnTop As Boolean = True, Optional ByVal blnLeft As Boolean = False)
    If blnLeft Then
        frmCenter.Move 0, (Screen.Height - frmCenter.ScaleHeight) / 2
    Else
        frmCenter.Move (Screen.Width - frmCenter.ScaleWidth) / 2, (Screen.Height - frmCenter.ScaleHeight) / 2
    End If
    
    If blnTop Then frmCenter.Move frmCenter.Left, 0
End Sub

Public Sub Sleep(ByVal lLoop As Long)
    Dim lCounter As Long
    
    For lCounter = 0 To lLoop
        DoEvents
    Next lCounter
End Sub
