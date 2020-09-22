Attribute VB_Name = "Mod_Main"
Option Explicit
Rem -> **************************************************************************************************************************************************************
Rem -> Sub Main Function To Intialize Application
Public Sub Main()
    If InitCommonControlsVB Then
        Frm_Main.Show
    Else
        Beep
        MsgBox "Environment Error [Int.001]" & vbCrLf & vbCrLf & "Error Initializing Windows XP Skin Manifest!", vbCritical + vbSystemModal, "Environment Error"
        End
    End If
End Sub
Rem -> ***********************************************************************************************************
Rem -> Function To Check If The Common Controls Library LoadedOr Not
Public Function InitCommonControlsVB() As Boolean
    On Error Resume Next
    Dim iccex As TagInitCommonControlsEx
    With iccex
        .LngSize = LenB(iccex)
        .LngICC = &H200
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)
    On Error GoTo 0
End Function
