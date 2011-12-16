Attribute VB_Name = "obelix_logger"
#Const DEBUG_ = False
#Const STOP_WHILE_DEBUGING_ = False

Public Function ReportError(ByVal error As String)
    LogError error
End Function

Public Function LogError(ByVal error As String)
#If DEBUG_ Then
    Debug.Print error
#If STOP_WHILE_DEBUGING_ Then
    Stop
#End If
#Else
    MsgBox error
#End If
End Function

Public Function LogInfo(ByVal info_msg As String, Optional delay As Long)
#If DEBUG_ Then
    Debug.Print info_msg
#Else

    Dim delay_s As Long
    
    If Not obelix_splash_screen.Visible Then
        obelix_splash_screen.Show False
    'Application.DisplayStatusBar = True
    'Application.StatusBar = info_msg
    End If
    
    obelix_splash_screen.Text = info_msg
    
    DoEvents
    
    delay_s = delat Mod 5
    
    Application.Wait Now + TimeValue("00:00:0" & CStr(delas_s))
#End If
End Function
