Attribute VB_Name = "MustException"
#Const DEBUG_ = True

Public Function ReportError(ByVal error As String)
#If DEBUG_ Then
    Debug.Print error
#Else
    MsgBox error
#End If
End Function
