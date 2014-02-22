Attribute VB_Name = "obelix_config"
'**
'* Gets a value that is related with the specified variable.
'*
'* <p>The worksheet whose code name is [config] will be used by this method and it must exists at
'* the current workbook.
'*
'* @param named_constant_var The variable that holds the configuration value.
'* @param encrypted_flag_offset The offset of the cell that contains a flag indicating if the configuration values
'*                              is encrypted.
'*
'* @return The value related with the specified variable or Empty if the configured value could not be retrieved.
'*
Public Function GetConfiguredValue(ByVal named_constant_var, Optional encrypted_flag_offset As Integer = -1) As String
    Dim config_value As String
    Dim is_encrypted As Boolean
    Dim crypt As Object
    Dim current_config_range As Range
    
    GetConfiguredValue = False
    
    Set crypt = Nothing
    Set current_config_range = config.Range(kConfigTopRightCell)
    
    If SearchFor(current_config_range, named_constant_var) Then
        config_value = current_config_range.Offset(, 1).value
        
        ' check if the value is encrypted
        If (encrypted_flag_offset > 0) Then
            If StrConv( _
                    current_config_range.Offset(0, encrypted_flag_offset).value, vbLowerCase _
                ) = "y" Then
                
                ' the value is encrypted, we need to instantiate the crypt class and decrypt the value.
                If crypt Is Nothing Then
                    If Not obelix_win.CreateRegFreeObject(obelix_consts.kCryptClassName, _
                        obelix_io.GetBinPathFor( _
                            obelix_consts.kdhRichClientDLLName _
                        ), crypt) Then
                        
                        Err.Raise vbObjectError + 513, _
                            "The " & obelix_consts.kCryptClassName & " could not be created."
                            
                        GoTo Finally
                    End If ' CreateRegFreeObject
                End If ' Is Nothing check
                
                config_value = crypt.Base64Dec(config_value)
            End If ' StrConv
        End If
        
        GoTo Finally
    End If ' Search for
    
    ' if the code reach here means taht the searched value was not found, so a exception
    ' must be raised.
    Err.Raise vbObjectError + 600, _
        Description:="Missing configuration for " & named_constant_var
        
Finally:
    GetConfiguredValue = config_value
    If Not crypt Is Nothing Then
        Set crypt = Nothing
    End If
    
    If Err.number <> 0 Then
        ' propagate the error to the caller
        Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

