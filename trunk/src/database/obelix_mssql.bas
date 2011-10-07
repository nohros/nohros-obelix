Attribute VB_Name = "obelix_mssql"
Option Explicit
Option Base 1

'**
'* A type that associates a sql command string with a ADODB connection that can be used to execute
'* the sql command.
Public Type SQLConnectionCommand
    SQLCommand As String
    SQLConnection As Object ' ADODB.Connection
End Type

' Executes a SQL command and store the data into a worksheet starting at the specified cell
'
' sql_command: The SQL command to execute.
' sql_connection: The string used to connects to the database.
' data_first_cell: A range representing the starting point where the returned data will be inserted.
' command_parms: The list of command parameters values. This paramters must be in the same order as the
'                specified on the SQL command. It must be nothing if the SQL command does not have parameters.
'
' History:
'   2011.01.19 - neylor.silva
'     Release
Public Function SQLQueryToRange(ByVal sql_command_text As String, ByVal sql_connection_string As String, ByRef data_first_cell As Range, ParamArray command_parms() As Variant) As Integer
    Dim sql_connection As ADODB.Connection
    Dim sql_command As ADODB.Command
    Dim sql_recordset As ADODB.Recordset
    Dim no_of_columns As Integer
    Dim rows As Variant
    
    On Error GoTo Catch
    
    Set sql_connection = New ADODB.Connection
    sql_connection.Provider = "sqloledb" ' If the user has excel then it have this provider.
    sql_connection.ConnectionString = sql_connection_string
    
    SQLQueryToRange2 sql_command_text, sql_connection, data_first_cell, command_parms
    
    GoTo Finally

Catch:
    ReportError "MustSQLServer->SQLiteQueryToRange: " & Err.Description

Finally:
    If Not sql_connection Is Nothing Then
        If sql_connection.State = adStateOpen Then _
            sql_connection.Close
        Set sql_connection = Nothing
    End If
End Function

' Executes a SQL command and store the data into a worksheet starting at the specified cell
'
' sql_command: The SQL command to execute.
' sql_connection: The connection to the SQL Server. It must be opened.
' data_first_cell: A range representing the starting point where the returned data will be inserted.
' command_parms: The list of command parameters values. This paramters must be in the same order as the
'                specified on the SQL command. It must be nothing if the SQL command does not have parameters.
'
' History:
'   2011.01.19 - neylor.silva
'     Release
Public Function SQLQueryToRange2(ByVal sql_command_text As String, ByVal sql_connection As Connection, ByRef data_first_cell As Range, ParamArray command_parms() As Variant) As Integer
    Dim sql_command As ADODB.Command
    Dim sql_recordset As ADODB.Recordset
    Dim no_of_columns As Integer
    Dim rows As Variant
    
    On Error GoTo Catch
        
    Set sql_command = New ADODB.Command
    sql_command.CommandText = sql_command_text
    
    sql_command.ActiveConnection = sql_connection
    Set sql_recordset = sql_command.Execute

    ' copy the data to the range
    data_first_cell.CopyFromRecordset sql_recordset
    
    Exit Function
    
Catch:
    ReportError "MustSQLServer->SQLiteQueryToRange: " & Err.Description
End Function

' Executes a SQL command and store the data into an in memory array
'
'
' History:
'   2011.01.19 - neylor.silva
'     Release
Public Function SQLQueryToArray(ByVal sql_command_text As String, ByVal sql_connection As Connection, ByRef rows() As Variant, ByRef fields() As String, ByRef fields_types() As String) As Boolean
    Dim sql_command As ADODB.Command
    Dim sql_recordset As ADODB.Recordset
    Dim no_of_columns As Integer
    Dim no_of_rows As Long
    Dim iterator_i As Long
    Dim current_field As ADODB.Field
    
    On Error GoTo Catch
        
    Set sql_command = New ADODB.Command
    sql_command.CommandText = sql_command_text
    
    sql_command.ActiveConnection = sql_connection
    Set sql_recordset = sql_command.Execute
    
    If Not sql_recordset.EOF Then
        ' get the field nams from the resulted data set
        no_of_columns = sql_recordset.fields.Count
        ReDim fields(no_of_columns)
        ReDim fields_types(no_of_columns)
        
        For iterator_i = 1 To no_of_columns
            Set current_field = sql_recordset.fields(iterator_i - 1)
            fields(iterator_i) = current_field.name
            fields_types(iterator_i) = DataTypeToSQLiteType(current_field.Type)
        Next iterator_i
        
        rows = sql_recordset.GetRows
        
        SQLQueryToArray = True
    Else
        SQLQueryToArray = False
    End If
    
    GoTo Finally
    
Catch:
    LogError "[obelix_mssql   SQLQueryToArray]   " & Err.Description
    
Finally:
    Set sql_recordset = Nothing
End Function

Public Function SetAllConnectionsTo(ByRef mssql_connection_commands() As SQLConnectionCommand, ByVal mssql_connection As Object)
    Dim iterator_i As Long
    Dim no_of_rows As Long
    
    no_of_rows = UBound(mssql_connection_commands)
    
    For iterator_i = 1 To no_of_rows
        Set mssql_connection_commands(iterator_i).SQLConnection = mssql_connection
    Next iterator_i
End Function
