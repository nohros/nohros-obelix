VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObelixSQLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

' Set to true to runs the debug version of the code
#Const DEBUG_ = False

Public sqlite_connection As Object 'dhRichClient3.cConnection'
Private sqlite_factory As Object 'dhRichClient3.cFactory'
Private Const kMaxStatements As Long = 5

Private Const kCreateIndexCommandMask As String = "create $1 index if not exists [$2] on [$3]($4)"

Private Sub Class_Initialize()
    obelix_win.CreateRegFreeObject _
        obelix_consts.kSQLiteConnectionClassName _
            , obelix_io.GetBinPathFor( _
                obelix_consts.kSQLiteFileName _
            ), _
            sqlite_connection
End Sub

Private Sub Class_Terminate()
    Set sqlite_connection = Nothing
    Set sqlite_factory = Nothing
End Sub

'**
'* Opens the specified database.
'* <p> If a previously opened SQLite exists, it connection will be closed before the new connection is open.
'*
'* @param sqlite_connectionstring The path for the SQLite database.
'*
'* @return True if the operation succeeds; otherwise, false.
Function OpenConnection(ByVal sqlite_connection_string As String) As Boolean
    On Error GoTo Catch
    
    CloseConnection

    OpenConnection = sqlite_connection.OpenDB(sqlite_connection_string, "")
    
    GoTo Finally

Catch:
    LogError "[ObelixSqLite   OpenConnection]   " & Err.Description
    
Finally:
End Function

'**
'* Open a in memory database by copying the data and structure from the specified database.
'* <p> The underlying connection object will be set to connection that is the result of the copy operation.
'*
'* @param sqlite_connection_string Connection string to the database to copy.
Public Function CopyFromDatabase(ByVal sqlite_connection_string As String) As Boolean
    Dim sqlite_connection_source As Object
    Dim result As Boolean
    
    If obelix_win.CreateRegFreeObject( _
        kSQLiteConnectionClassName _
        , obelix_io.GetBinPathFor(obelix_consts.kdhRichClientDLLName) _
        , sqlite_connection_source) Then
        
        sqlite_connection_source.OpenDB sqlite_connection_string, ""
        Set sqlite_connection = sqlite_connection_source.CopyDatabase(":memory:")
    End If
    
    result = True
    
    GoTo Finally
    
Catch:
    result = False
    LogError Err.Description
    
Finally:
    CopyFromDatabase = result
End Function

'**
'* Closes a sqlite connection.
Function CloseConnection()
    ' the connection is auto closed
    'sqlite_connection.Close
End Function

'**
'* Executes a SQLite statement. Statement should be non-query but this is not necessary.
'* <p> The connection must be already opened.
'*
'* @param The statements to execute.
Public Function ExecSQLiteQuery(ByVal sql_command As String, Optional ByRef affected_rows As Long, Optional increment_affected_rows = False) As Boolean
    Dim sqlite_command As Object ' dhRichClient.cCommand
    Dim result As Boolean
    
    result = sqlite_connection.Execute(sql_command)
    
    affected_rows = IIf(increment_affected_rows, affected_rows, 0) + sqlite_connection.AffectedRows
    
    GoTo Finally
    
Catch:
    result = False
    LogError "[ObelixSQLite   ExecSQLiteQuery]   " + Err.Description & vbCrLf & _
        "SQL command: " & sql_command
    
Finally:
    ExecSQLiteQuery = result
End Function

'** Executes many SQL commands at once. The commands should be non-query.
'*
'* @param sqlite_commands A list of SQL commands separated by comma.
'* @param sql_connection The connection used to execute the SQL commands.
Public Function SQLiteQueryBatch(ByVal sqlite_command As String) As Boolean
    SQLiteQueryBatch = False

    sqlite_connection.BatchExecute sqlite_command

    SQLiteQueryBatch = True

    Exit Function
    
Catch:
    LogError "[ObelixSQLite   SQLiteQueryBatch]   " + Err.Description
End Function

' Refresh a table by cleaning up the data set and insert a new one
' table_name: The name of the table to refresh
'
' data_range: A spreedsheet range containing the data to refresh. The first row of the
'             range must contains the names of the columns that must be equals
'             to the names of the column defined for the specified table.
'
Public Function TableFromRange(ByVal table_name As String, ByVal data_range As Range) As Boolean
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
    
    TableFromRange = False ' Pesimist. false until true
    
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
    
    If Not ExecSQLiteQuery(sql_command) Then
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
        
        If Not ExecSQLiteQuery(sql_command) Then
            ReportError "RefreshTable->Error inserting data at row " + i
            Exit Function
        End If
    Next i
    
    TableFromRange = True
    
    Exit Function
    
Catch:
    ReportError "MustSQLite->RefreshTable " + Err.Description
    ReportError "MustSQLite->RefreshTable->SQL command:" & sql_command
End Function

' Executes a SQL command and store the data into a worksheet starting at the specified cell
'
' sqlite_command: The SQL command to execute.
' sql_connection: The SQLite connectione used to execute the command. The connection state must be open.
' data_first_cell: A range representing the starting point where the returned data will be inserted.
'
' History:
'   2011.01.01 - neylor.silva
'     Release
Public Function SQLiteQueryToRange(ByVal sql_command As String, ByRef data_first_cell As Range, Optional display_headers As Boolean = False) As Integer
    Dim sqlite_recordset As Object 'dhRichClient.cCommand
    Dim rows() As Variant 'GetRows()
    Dim in_memory_data As Variant
    Dim no_of_rows As Long
    Dim no_of_columns As Long
    Dim i As Long
    Dim j As Long
    
    On Error GoTo Catch
    
    Set sqlite_recordset = sqlite_connection.OpenRecordset(sql_command)
    
    ' We need to know the number of rows in order to create an array that can be used
    ' to transfer the data to an worksheet(this is the fast way). We cannot use the Step
    ' method because this imply Redim Preserve and this cannot be used too. The rows of
    ' the array must be the first dimension and Redim Preserve cannot redimenssion this
    ' type of array.
    If sqlite_recordset.EOF Then
        GoTo Finally
    End If
    
    in_memory_data = IIf(display_headers, sqlite_recordset.GetRowsWithHeaders(transposed:=True), sqlite_recordset.GetRows(transposed:=True))
    
    ' The returned array is a;ways zero based
    no_of_rows = UBound(in_memory_data) + 1
    no_of_columns = UBound(in_memory_data, 2) + 1
    
    ' write the data to the sheet starting at the specified cell
    data_first_cell.Resize(no_of_rows, no_of_columns) = in_memory_data
    
    SQLiteQueryToRange = no_of_rows
    
    GoTo Finally
    
Catch:
    ReportError "MustSQLite->SQLiteQueryToRange " + Err.Description
    
Finally:
    Set sqlite_recordset = Nothing
End Function

Public Function SQLiteQueryToArray(ByVal sql_command_text As String, ByRef rows() As Variant, ByRef fields() As String, ByRef fields_types() As String, Optional transposed As Boolean = False) As Boolean
    Dim sqlite_recordset As Object
    Dim no_of_columns As Integer
    Dim no_of_rows As Long
    Dim iterator_i As Long
    Dim current_field As Object
    Dim result As Boolean
    
    On Error GoTo Catch
        
    Set sqlite_recordset = sqlite_connection.OpenRecordset(sql_command_text)
    
    ' We need to know the number of rows in order to create an array that can be used
    ' to transfer the data to an worksheet(this is the fast way). We cannot use the Step
    ' method because this imply Redim Preserve and this cannot be used too. The rows of
    ' the array must be the first dimension and Redim Preserve cannot redimenssion this
    ' type of array.
    If sqlite_recordset.EOF Then
        GoTo Finally
    End If
    
    rows = sqlite_recordset.GetRows(transposed:=transposed)
    
    ' The returned array is always zero based
    If transposed Then
        no_of_rows = UBound(rows) + 1
        no_of_columns = UBound(rows, 2) + 1
    Else
        no_of_rows = UBound(rows, 2) + 1
        no_of_columns = UBound(rows) + 1
    End If
    
    ReDim fields(no_of_columns)
    ReDim fields_types(no_of_columns)
    
    For iterator_i = 1 To no_of_columns
        Set current_field = sqlite_recordset.fields(iterator_i - 1)
        fields(iterator_i) = current_field.name
        fields_types(iterator_i) = current_field.OriginalDataType
    Next iterator_i
    
    GoTo Finally
    
Catch:
    result = False
    LogError "[ObelixSQLite   SQLiteQueryToArray] " + Err.Description
    
Finally:
    Set sqlite_recordset = Nothing
    SQLiteQueryToArray = no_of_rows > 0
End Function

'**
'* Copies the contents of the specifed array to the specified SQLite table.
'* <p>The number of columns of the data array must be equals to the number of the columns of
'* the columns_names array.
'*
'* @param table_name The name of the destination table.
'* @param data An two dimensional array containing the data to be copied. The first subscript of the array
'*             identifies the field and the second identifies the record number.
'* @param column_names An array containing the names of the columns of the SQLite table.
'* @param columns_types The types of the columns. This parameter is used to create the destination table if it does
'*                      not exists into the database. It could be omitted if the table already exists. If a type was not specified
'*                      the related colum will be created using the "text" type.
Public Function TableFromArray(ByVal table_name As String, ByRef data() As Variant, ByRef columns_names() As String, ByRef columns_types() As String) As Boolean
    Dim sqlite_command_mask As String
    Dim sqlite_command As String
    Dim columns_mask As String
    Dim no_of_columns As Integer
    Dim no_of_rows As Long
    Dim iterator_i As Long
    Dim iterator_j As Long
    Dim bound_offset As Long
    Dim no_of_statements As Long
    Dim cs_data As String
    
    Dim result As Boolean
    
    On Error GoTo Catch
       
    ' An value that will be used to fix the default array lower bound subscript. The main objective
    ' of this method is to copy data from arrays ceated by the Recordset.GetRows() method. That method
    ' always make a array with 0 as lower bound subscript. So by default we will assume that the lower
    ' bound of the data array is zero. If data array is not zero-based we need to adjust the bound_offset
    ' in order to perform loops correctly
    bound_offset = IIf(LBound(data) = 0, 1, 0)
    
    no_of_columns = UBound(data) + bound_offset
    no_of_rows = UBound(data, 2) + bound_offset
    
    If no_of_columns <> UBound(columns_names) Then
        LogError "[obelix_sqlite   TableFromArray]   " & "obelix_sqlite_err_no_of_columns"
        GoTo Finally
    End If

    result = CreateTable(table_name, columns_names, columns_types)
    If Not result Then
        'TODO? log the error
        GoTo Finally
    End If
    
    sqlite_command_mask = "insert into $1($2) values($3)"
    
    ' build a string containing a mask for the columns
    For iterator_i = 1 To no_of_columns
        columns_mask = columns_mask & _
            obelix_sql.EscapeWithTralingComma( _
                columns_names(iterator_i) _
            )
    Next iterator_i
    
    ' remove the trailing comma
    columns_mask = RemoveTrailing(columns_mask)
    
    ' the sql command maks will be something like
    '    insert into table_name(column1, column2, ...) values($1, $2, ...)
    sqlite_command_mask = FORMATARTEXTO(sqlite_command_mask, table_name, columns_mask, "$1")
    
    ReDim column_data(no_of_columns)
    
    ' build a string containing a mask for the values
    For iterator_i = 1 To no_of_rows
        cs_data = ""
        For iterator_j = 1 To no_of_columns
            cs_data = cs_data & _
                obelix_sql.QuoteWithTralingComma( _
                    data(iterator_j - bound_offset, iterator_i - bound_offset) _
                )
        Next iterator_j
        
        sqlite_command = sqlite_command & FORMATARTEXTO(sqlite_command_mask, RemoveTrailing(cs_data)) & ";"
        
        no_of_statements = no_of_statements + 1
        If no_of_statements > kMaxStatements Then
            result = Me.ExecSQLiteQuery(sqlite_command)
            no_of_statements = 0
            sqlite_command = ""
        End If
        
        DoEvents
    Next iterator_i
    
    If no_of_statements > 0 Then
        result = Me.ExecSQLiteQuery(sqlite_command)
    End If
    
    GoTo Finally
    
Catch:
    result = False
    LogError "[obelix_sqlite   TableFromArray]   " & Err.Description
    
Finally:
    TableFromArray = result
End Function

'**
'* Creates a table if it does not exist by using the specified table name, columns names and types.
'*
'* @param table_name The name of the table to create.
'* @param columns_names An array containing the names of the table columns.
'* @param colmns_types An array containing the type name of each columns. If a type was not specified
'* the columns will be created as a text column.
'*
'* @return true if the table was succeesfully created or already exists; otherwise, false.
Public Function CreateTable(ByVal table_name As String, ByRef columns_names() As String, ByRef columns_types() As String)
    Dim sqlite_command As String
    Dim columns_mask As String
    Dim iterator_i As Long
    Dim no_of_columns As Long
    Dim no_of_types As Long
    Dim result As Boolean
    
    On Error GoTo Catch
    
    no_of_columns = UBound(columns_names)
    no_of_types = UBound(columns_types)
    
    ' creates the table to hold the data and build the insert command
    sqlite_command = "create table if not exists $1($2)"
    
    ' build a string containing a mask for the values
    For iterator_i = 1 To no_of_columns
        columns_mask = columns_mask & _
            obelix_sql.EscapeWithTralingComma( _
                columns_names(iterator_i) & " " & IIf(iterator_i <= no_of_types, columns_types(iterator_i), "text") _
            )
    Next iterator_i
    
    columns_mask = RemoveTrailing(columns_mask)
    
    sqlite_command = FORMATARTEXTO(sqlite_command, table_name, columns_mask)
    
    result = ExecSQLiteQuery(sqlite_command)

    GoTo Finally
    
Catch:
    result = False
    LogError "[ObelixSQLite   CreateTable]   " & Err.Description
    
Finally:
    CreateTable = result

End Function

'**
'* Create a index on the specifed table by using the speicifed name and list of columns.
'* <p> If another index with the same name already exists, then this command becomes a no-op.
'*
'* @param table_name The name of the table to be indexed.
'* @param index_name The name of the index. Index names must be unique within a table.
'* @param field_list The list of columns on which the index is based.
Public Function CreateIndexOn(ByVal table_name As String, ByVal index_name As String, ByVal field_list As String, Optional is_unique As Boolean = False) As Boolean
    Dim create_index_command As String
    Dim result As Boolean
    
    On Error GoTo Catch
    
    result = False
    
    create_index_command = FORMATARTEXTO(kCreateIndexCommandMask, IIf(is_unique, "unique", ""), index_name, table_name, field_list)
    
    result = Me.ExecSQLiteQuery(create_index_command)
    
    GoTo Finally
    
Catch:
    LogError Err.Description
    
    result = False
Finally:
    CreateIndexOn = result
End Function

Public Property Get Connection() As Object
    Set Connection = sqlite_connection
End Property

