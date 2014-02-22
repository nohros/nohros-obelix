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
    ReportError Err

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

    Set sql_recordset = ExecuteSQLQuery(sql_command_text, sql_connection, _
        command_parms)

    ' copy the data to the range
    data_first_cell.CopyFromRecordset sql_recordset
    
    GoTo Finally
    
Catch:
    LogError Err

Finally:
    CloseRecordset sql_recordset
    If Err.number <> 0 Then Throw Err
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
Public Function SQLProcedureToRange2(ByVal sql_command_text As String, ByVal sql_connection As Connection, ByRef data_first_cell As Range, ParamArray command_parms() As Variant) As Integer
    ExecuteSQLQuery "SET NOCOUNT ON", sql_connection
    SQLQueryToRange2 sql_command_text, sql_connection, data_first_cell, command_parms
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
    LogError Err
    
Finally:
    Set sql_recordset = Nothing
    
If Err.number <> 0 Then Throw Err
End Function

Public Function SqlQueryScalar( _
    ByVal sql_command_text As String, _
    ByVal sql_connection As Connection, _
    ParamArray command_parms() As Variant) As Variant

    Dim sql_recordset As ADODB.Recordset
    Dim iterator_i As Long
    Dim current_field As ADODB.Field
    
    On Error GoTo Catch

    Set sql_recordset = ExecuteSQLQuery(sql_command_text, sql_connection, command_parms)
    
    If Not sql_recordset.EOF Then
        SqlQueryScalar = sql_recordset(0)
    Else
        SqlQueryScalar = Nothing
    End If
    
    GoTo Finally
    
Catch:
    LogError Err
    
Finally:
    CloseRecordset sql_recordset
    
    If Err.number <> 0 Then Throw Err
End Function

Private Function ExecuteSQLQuery( _
    ByVal sql_command_text As String, _
    ByVal sql_connection As Connection, _
    ParamArray command_parms() As Variant) As Recordset

    Dim sql_command As ADODB.Command
    
    On Error GoTo Catch

    Set sql_command = New ADODB.Command
    sql_command.CommandText = FormatarTexto(sql_command_text, command_parms)
    
    If sql_connection.State = adStateClosed Then
        sql_connection.Open
    End If
    
    sql_command.ActiveConnection = sql_connection
    Set ExecuteSQLQuery = sql_command.Execute
    
    GoTo Finally
    
Catch:
    LogError Err
    Throw Err
    
Finally:
End Function

Public Function SetAllConnectionsTo(ByRef mssql_connection_commands() As SQLConnectionCommand, ByVal mssql_connection As Object)
    Dim iterator_i As Long
    Dim no_of_rows As Long
    
    no_of_rows = UBound(mssql_connection_commands)
    
    For iterator_i = 1 To no_of_rows
        Set mssql_connection_commands(iterator_i).SQLConnection = mssql_connection
    Next iterator_i
End Function

' Closes tge given SQL connection object.
'
' @remarks If the given sql connection is not opened or is |Nothing| this
'          method does not perform any operation. After closed the given
'          |sql_connection| object will be assigned to |Nothing|
'
' History:
'   2014.01.29 - neylor.silva
'     Release
Public Function CloseConnection(ByRef sql_connection As Connection)
    If Not sql_connection Is Nothing Then
        If sql_connection.State = adStateOpen Then
            sql_connection.Close
        End If
        Set sql_connection = Nothing
    End If
End Function

Public Function CloseRecordset(ByRef sql_recordset As Recordset)
    If Not sql_recordset Is Nothing Then
        If sql_recordset.State = adStateOpen Then
            sql_recordset.Close
        End If
        Set sql_recordset = Nothing
    End If
End Function

' Create a new Connection object and associates the given connection string to it
'
' @connection_string The string that contains the information used to connect to
'                    a database server.
Public Function CreateConnection(ByVal connection_string As String) As Connection
    On Error GoTo Catch:
    
    Set CreateConnection = New Connection
    CreateConnection.ConnectionString = connection_string
    
    GoTo Finally
    
Catch:
    LogError Err
    
Finally:
End Function

Public Function OpenConnection(ByVal connection_string As String) As Connection
    On Error GoTo Catch:
    
    Set OpenConnection = New Connection
    OpenConnection.Open connection_string

    GoTo Finally
    
Catch:
    LogError Err
    Throw Err
    
Finally:
End Function


Public Function RangeToSQL(ByVal table_name As String, _
    ByVal data_range As Range, _
    ByVal sql_connection As Connection) As Boolean
    
    Dim in_memory_data As Variant
    Dim in_memory_column_names As Variant
    Dim in_memory_column_types As Variant
    Dim no_of_columns As Long
    Dim no_of_rows As Long
    Dim first_element As Range
    Dim sql_command As String
    Dim sql_insert_command As String
    Dim i As Long
    Dim j As Long
    
    RangeToSQL = False ' Pesimist. false until true
    
    On Error GoTo Catch
    
    no_of_columns = data_range.Columns.Count
    no_of_rows = data_range.rows.Count
        
    ' reads the data into memory to avoid communication overhead
    in_memory_column_names = data_range.Resize(1) ' gets the column names
    
    ' build the insert command
    sql_insert_command = "insert into " & table_name & "("
    For i = 1 To no_of_columns
        sql_insert_command = sql_insert_command & "[" & in_memory_column_names(1, i) & "],"
    Next i
    sql_insert_command = RemoveTrailing(sql_insert_command) & ") values("
    
    ' load the data from the given range
    If no_of_rows = 1 Then
        in_memory_data = Array(0) ' no data
        no_of_rows = 0 ' adjust the number of rows
    Else
        no_of_rows = no_of_rows - 1 ' remove the header from the row count
        in_memory_data = data_range.Resize(no_of_rows).Offset(1)
    End If
    
    ' fill the table with data.
    For i = 1 To no_of_rows
        sql_command = sql_insert_command
        For j = 1 To no_of_columns
            sql_command = sql_command & QuoteWithTralingComma(in_memory_data(i, j))
        Next j
        sql_command = RemoveTrailing(sql_command) + ")"
        
        If Not ExecuteSQLQuery(sql_command, sql_connection) Then
            Err.Raise vbObjectError + 512, _
                "obelix_mssql.RangeToSQL", _
                FormatarTexto("The row $1 fails to be imported", i), _
                "", ""
        End If
    Next i
    
    RangeToSQL = True
    
    GoTo Finally
    
Catch:
    LogError Err
    Throw Err
    
Finally:
End Function
