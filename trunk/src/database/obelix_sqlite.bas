Attribute VB_Name = "obelix_sqlite"
Option Explicit
Option Base 1

Public Function RegisterLitex(ByVal regsvrex_path As String) As Boolean
'    Dim sqlite_conn As Object
'    Dim sqlite_version As String
'
'    On Error GoTo Catch
'
'    ' register the Litex OCX
'    Shell regsvrex_path & " /c " & SQLITE_DLL_PATH
'
'    ' check if it is successfully registered
'    On Error Resume Next
'    Set sqlite_conn = CreateObject("Litex.Object")
'    RegisterLitex = (Not sqlite_conn Is Nothing)
'
'    GoTo Finally
'
'Catch:
'    RegisterLitex = False
'
'Finally:
'    If Not sqlite_conn Is Nothing Then _
'        Set sqlite_conn = Nothing
End Function

' Executes a SQL command and store the data into a worksheet starting at the specified cell
'
' sql_command: The SQL command to execute.
' sql_connection: The SQLite connectione used to execute the command. The connection state must be open.
' data_first_cell: A range representing the starting point where the returned data will be inserted.
' command_parms: The list of command parameters values. This paramters must be in the same order as the
'                specified on the SQL command. It must be nothing if the SQL command does not have parameters.
'
' History:
'   2011.01.01 - neylor.silva
'     Release
Public Function SQLiteQueryToRange(ByVal sql_command As String, ByRef sql_connection As Object, ByRef data_first_cell As Range, ParamArray command_parms() As Variant) As Integer
    Dim sql_statement As Object 'Litex.LiteStatement
    Dim rows As Object 'Litex.LiteRows
    Dim in_memory_data As Variant
    Dim no_of_rows As Long
    Dim no_of_columns As Long
    Dim i As Long
    Dim j As Long
    
    On Error GoTo Catch
    
    Set sql_statement = sql_connection.Prepare(sql_command)
    
    If Not IsEmpty(command_parms) Then
        For i = 1 To UBound(command_parms)
            sql_statement.BindParameter i - 1, command_parms(i - 1)
        Next i
    End If

    no_of_columns = sql_statement.ColumnCount
    ' We need to know the number of rows in order to create an array that can be used
    ' to transfer the data to an worksheet(this is the fast way). We cannot use the Step
    ' method because this imply Redim Preserve and this cannot be used too. The rows of
    ' the array must be the first dimension and Redim Preserve cannot redimenssion this
    ' type of array.
    Set rows = sql_statement.rows(True)
    no_of_rows = rows.Count
    
    If no_of_rows = 0 Then GoTo Finally
    
    
    ' This array will be used to transfer the returned data to an worksheet in a fast way.
    ReDim in_memory_data(no_of_rows, no_of_columns)
    For i = 1 To no_of_rows
        For j = 1 To no_of_columns
            in_memory_data(i, j) = rows(i - 1)(j - 1) ' rows array is zero-based
        Next j
    Next i
    
    ' write the data to the sheet starting at the specified cell
    data_first_cell.Resize(no_of_rows, no_of_columns) = in_memory_data
    
    SQLiteQueryToRange = no_of_rows
    
    GoTo Finally
    
Catch:
    ReportError "MustSQLite->SQLiteQueryToRange " + Err.Description
    
Finally:
    If Not sql_statement Is Nothing Then _
        sql_statement.Close
End Function

' Executes many SQL commands at once. The commandS cannot return rows.
'
' sql_commands: A list of SQL commands separated by comma.
' sql_connection: The connection used to execute the SQL commands.
'
' History:
'   2010.01.12 - neylor.silva
'     Release
Public Function SQLiteQueryBatch(ByVal sql_command As String, ByRef sql_connection As Object) As Boolean
    SQLiteQueryBatch = False
        
    sql_connection.BatchExecute sql_command
    
    SQLiteQueryBatch = True
    
    Exit Function
    
Catch:
    ReportError "MustSQLite->SQLiteQueryBatch " + Err.Description
    ReportError "MustSQLite->RefreshTable->SQL command:" & sql_command
End Function

'**
'* Call the SQLQueryToArray function and pass the resulting array to the
'* SQLiteTableFromArray method.
'*
Public Function SQLQueryToSQLite(ByVal mssql_command As String, ByVal mssql_connection As ADODB.Connection, _
    ByVal sqlite_table_name As String, ByVal obelix_sqlite As ObelixSQLite) As Boolean
    
    Dim data_rows() As Variant
    Dim data_fields() As String
    Dim data_fields_types() As String
    Dim result As Boolean
    
    On Error GoTo Catch
    
    result = SQLQueryToArray(mssql_command, mssql_connection, data_rows, data_fields, data_fields_types)
    If result Then
        result = obelix_sqlite.TableFromArray(sqlite_table_name, data_rows, data_fields, data_fields_types)
    End If
    
    GoTo Finally
    
Catch:
    result = False
    LogError Err.Description
Resume
Finally:
    SQLQueryToSQLite = result
End Function
