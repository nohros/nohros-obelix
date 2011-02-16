Attribute VB_Name = "MustSQLServer"
' Copyright (c) 2010 Nohros Systems Inc.
' Copyright (c) 2003 Dermot Balson.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the "Software"), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
' to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.
'
'
' This...


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
    sql_connection.Provider = "sqloledb" ' If the user has excel then  it have this provider.
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
