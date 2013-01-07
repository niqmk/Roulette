Attribute VB_Name = "mdMain"
Option Explicit

Public BLogin As Boolean

Sub Main()
    SetInitial
    
    fSplash.Show vbModal
    fLogin.Show vbModal
    
    If BLogin Then
        fMain.Show
    End If
End Sub

Private Sub SetInitial()
    If Right(App.Path, 1) = "\" Then
        mdGeneral.SPath = App.Path
    Else
        mdGeneral.SPath = App.Path & "\"
    End If
End Sub
