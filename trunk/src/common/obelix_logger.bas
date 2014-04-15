Attribute VB_Name = "obelix_logger"
#Const DEBUG_ = False
#Const STOP_WHILE_DEBUGING_ = False

Public Function ReportError(ByVal error As ErrObject)
    MsgBox error.Description
End Function

Public Function LogError(ByVal error As ErrObject)
#If DEBUG_ Then
    Debug.Print error.Description
#If STOP_WHILE_DEBUGING_ Then
    Stop
#End If
#Else
#End If
End Function

Public Function Throw(ByVal Err As ErrObject)
    Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function Finally(ByVal Err As ErrObject)
  If Err.Number <> 0 Then
    On Error GoTo 0
    Throw Err
  End If
End Function
