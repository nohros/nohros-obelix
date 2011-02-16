Attribute VB_Name = "MustSQLite"
Option Explicit
Option Base 1

' Set to true to runs the debug version of the code
#Const DEBUG_ = True

Public Type SQLiteCommand
    CommandText As String
End Type

Public Type SQLiteConecction
    Connection As Object 'Litex.LiteConnection
    Command As SQLiteCommand
End Type

Private in_memory_sqlite_database As Object

Public Sub Configure(ByVal current_workbook As Workbook)
    Dim fs As FileSystemObject
    Dim regsvrex_path As String
    Dim data As Variant
    Dim sql As String
    Dim i As Long
    Dim j As Long
    Dim size As Long
   
    Set fs = New FileSystemObject
    
    ' set the paths
    BASE_PATH = current_workbook.path ' The spreedsheet base path
    SQLITE_DLL_PATH = fs.BuildPath(BASE_PATH, kSQLiteFileName)
    SQLITE_DB_PATH = fs.BuildPath(BASE_PATH, kSQLiteDataBase)
    regsvrex_path = fs.BuildPath(BASE_PATH, kRegSvrExFileName)
        
    ' ensure the existence of the sqlite dll
    If Not fs.FileExists(SQLITE_DLL_PATH) Then _
        WriteEmbedToDisk current_workbook, kSQLiteFileName, BASE_PATH
    
    ' ensure the existence of the database file
    If Not fs.FileExists(SQLITE_DB_PATH) Then _
        WriteEmbedToDisk current_workbook, kSQLiteDataBase, BASE_PATH
    
    ' ensure the existence of the reggie software
    If Not fs.FileExists(regsvrex_path) Then _
        WriteEmbedToDisk current_workbook, kRegSvrExFileName, BASE_PATH
        
    ' register the Litex SQLite OCX wrapper
    RegisterLitex regsvrex_path
    
Finally:
End Sub
    
Public Function RegisterLitex(ByVal regsvrex_path As String) As Boolean
    Dim sqlite_conn As Object
    Dim sqlite_version As String
    
    On Error GoTo Catch
    
    ' register the Litex OCX
    Shell regsvrex_path & " /c " & SQLITE_DLL_PATH
        
    ' check if it is successfully registered
    On Error Resume Next
    Set sqlite_conn = CreateObject("Litex.Object")
    RegisterLitex = (Not sqlite_conn Is Nothing)
    
    GoTo Finally
    
Catch:
    RegisterLitex = False
    
Finally:
    If Not sqlite_conn Is Nothing Then _
        Set sqlite_conn = Nothing
End Function

Public Sub CloseInMemoryDatabase()
    If Not in_memory_sqlite_database Is Nothing Then
        in_memory_sqlite_database.Close
        Set in_memory_sqlite_database = Nothing
    End If
End Sub

Public Function GetInMemoryDatabase() As Object
    On Error GoTo Catch
    
    If in_memory_sqlite_database Is Nothing Then
        ' creates a in-memory database connection ...
        Set in_memory_sqlite_database = CreateObject("Litex.LiteConnection")
        in_memory_sqlite_database.OpenInMemory
        
        '... and fill it with some data.
        If Not RefreshDatabase(in_memory_sqlite_database) Then  ' This function must be supplied by the application
            in_memory_sqlite_database.Close
            Set in_memory_sqlite_database = Nothing
            ReportError "MustSQLite->GetInMemoryDatabase: The in-memory database could not be refreshed."
        End If
    End If
    
    Set GetInMemoryDatabase = in_memory_sqlite_database
    
    Exit Function
    
Catch:
    ' On error we need to invalidate the in_memory database
    CloseInMemoryDatabase
    
    ReportError "GetInMemoryDatabase->" + Err.Description
End Function

' Removes the last character from a string. It is usually used to remove the last
' comma from a SQL command when this command is build inside a loop.
'
' data_with_trailing_char: The string to remove the last character.
Public Function RemoveTrailing(ByVal str As String)
    RemoveTrailing = Mid(str, 1, Len(str) - 1)
End Function

' Quotes a string.
Public Function Quote(ByVal str As String)
    Quote = "'" & str & "'"
End Function

' Quotes a string and add a comma at the end.
Public Function QuoteWithTralingComma(ByVal str As String)
    QuoteWithTralingComma = "'" & str & "',"
End Function

' Refresh a table by cleaning up the data set and insert a new one
' table_name: The name of the tbale to refresh
'
' data_range: A spreedsheet range containing the data to refresh. The first row of the
'             range must contains the names of the columns that must be equals
'             to the names of the column defined for the specified table.
'
Public Function RefreshTable(ByVal table_name As String, ByVal data_range As Range, ByRef sql_connection As Object) As Boolean
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
    
    RefreshTable = False ' Pesimist. false until true
    
    On Error GoTo Catch
    
    no_of_columns = data_range.Columns.Count
    no_of_rows = data_range.rows.Count
    
    'Set first_element = data_range.rows(1) 'TODO: Check if this could be removed from here. No use
    
    ' reads the data into memory to avoid communication overhead
    in_memory_column_names = data_range.Resize(1) ' gets the column names
    in_memory_column_types = data_range.Resize(2).Offset(1)
    
    ' creates the table to hold the data and build the insert command
    sql_command = "create table if not exists " & table_name & "("
    sql_insert_command = "insert into " & table_name & "("
    For i = 1 To no_of_columns
        sql_command = sql_command & in_memory_column_names(1, i) & " " & in_memory_column_types(1, i) & ","
        sql_insert_command = sql_insert_command & in_memory_column_names(1, i) & ","
    Next i
    sql_command = RemoveTrailing(sql_command) & ")"
    sql_insert_command = RemoveTrailing(sql_insert_command) & ") values("
    
    If Not ExecSQLiteQuery(sql_command, sql_connection) Then
        ReportError "RefreshTable->The table " & table_name & "could not be created."
        Exit Function
    End If
    
    ' Load the data into the newly created table.
    If no_of_rows < 3 Then
        in_memory_data = Array(0) ' no data
        no_of_rows = 0 ' adjust the number of rows
    Else
        no_of_rows = no_of_rows - 2 ' remove the header from the row count
        in_memory_data = data_range.Resize(no_of_rows).Offset(2) ' gets the data
    End If
    
    ' fill the table with data.
    For i = 1 To no_of_rows
        sql_command = sql_insert_command
        For j = 1 To no_of_columns
            sql_command = sql_command & QuoteWithTralingComma(in_memory_data(i, j))
        Next j
        sql_command = RemoveTrailing(sql_command) + ")"
        
        If Not ExecSQLiteQuery(sql_command, sql_connection) Then
            ReportError "RefreshTable->Error inserting data at row " + i
            Exit Function
        End If
    Next i
    
    RefreshTable = True
    
    Exit Function
    
Catch:
    ReportError "MustSQLite->RefreshTable " + Err.Description
    ReportError "MustSQLite->RefreshTable->SQL command:" & sql_command
End Function

Public Function ExecSQLiteQuery(ByVal sql_command As String, ByRef sql_connection As Object)
    Dim sql_statement As Object ' Litex.LiteStatement

    ExecSQLiteQuery = False
    
    Set sql_statement = sql_connection.Prepare(sql_command)
    sql_statement.Execute
    
    ExecSQLiteQuery = True
    
    GoTo Finally
    
Catch:
    ReportError "MustSQLite->ExecSQLiteQuery " + Err.Description
    
Finally:
    If Not sql_statement Is Nothing Then _
        sql_statement.Close
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
