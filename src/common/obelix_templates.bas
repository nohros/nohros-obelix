Attribute VB_Name = "obelix_templates"
Option Explicit
Option Base 1

Public kConstIsSet As Boolean

Public Const kMaxRows As Long = 65536
Public Const kMaxColumns As Long = 256

Public Type ClearContentFlags
    ClearConditionalFormats As Boolean
    ClearFormats As Boolean
End Type

'**
'* Refresh the in memory database. This method is called by the SQLite module
'* when it needs to refresh the in memory database.
'*
'* @param sql_connection: The connection to the in memory database.
'*
Public Function RefreshDatabase(ByRef sql_connection As Object) As Boolean
    Dim data_range As Range
    Dim first_data_row As Range
    Dim no_of_rows As Long
    Dim no_of_columns As Long
    
    On Error GoTo Catch
    
    RefreshDatabase = False
 
    EnsureConstants
 
    ' The code that do the refresh goes here
    '
    ' It is something like that:
    '
    '  If GetDataRange(GetConfiguredValue(kDataHeaderFirstCell), kDataColumnsCount, data_range, data) Then
    '     RefreshTable kCargaTableName, data_range, sql_connection
    '  Else
    '      ReportError "[ObelixTemplates RefreshDatabase]: Could not retrieve the address range related with the table [TABLE]"
    '      Goto Finally
    '  End If
    
    Exit Function
    
Catch:
    ReportError "[ObelixTemplates RefreshDatabase]   " & Err.Description
    
' Propagate the error to the caller.
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
Finally:
End Function

'**
'* Gets a sheet range that can be used to refresh a SQLite database.
'*
'* @param sql_connection: The connection to the in memory database.
'*
'* @return True if a valid range could be retrievedç otherwise, false.
'*         <p> A range is valid when the first row contains the names of the columns and the
'*         others contains the data. At least one row must have data and none of the columns
'*         can have a empty value.
'*
Private Function GetDataRange(ByVal first_data_cell_address, ByVal no_of_columns As Long, ByRef data_range As Range, ByRef data_sheet As Worksheet) As Boolean
    Dim first_data_row As Range
    Dim single_cell As Range
    Dim temp_range As Range
    Dim no_of_rows As Long
        
    Set first_data_row = data_sheet.Range(first_data_cell_address).Resize(, no_of_columns)
    no_of_rows = first_data_row.End(xlDown).Row
    
    Set data_range = first_data_row.Resize(no_of_rows, no_of_columns)
    
    ' check if the all column names are supplied
    For Each single_cell In first_data_row.Cells
        If single_cell.value = Empty Then
            GetDataRange = False
            Exit Function
        End If
    Next single_cell
    
    ' check if the data range is empty.
    If no_of_rows = kMaxRows Then
        Set temp_range = data_range
        Set data_range = first_data_row
        Set first_data_row = first_data_row.Offset(1) ' first row after the header
        
        For Each single_cell In first_data_row.Cells
            If single_cell.value <> Empty Then
                Set data_range = temp_range
                Exit For
            End If
        Next single_cell
    End If
    
    GetDataRange = True
End Function

'**
'* Clear the contents of the specified range. The vertical and horizontal size of the content is dynamically calculated.
'*
'* @param first_data_cell The initial cell of the range.
'*
Public Function ClearRangeContentDynamic(ByVal first_data_cell As Range, ByRef clear_flags As ClearContentFlags) As String
    Dim no_of_columns As Long
    Dim no_of_rows As Long
    
    ' Get the position of the last horizontal.
    no_of_columns = first_data_cell.End(xlToRight).Column - first_data_cell.Column + 1
    no_of_rows = first_data_cell.End(xlDown).Row - first_data_cell.Row + 1

    ClearRangeContentDynamic = ClearRangeContentSize(first_data_cell, no_of_columns, no_of_rows, clear_flags)
End Function

'**
'* Clear the contents of the specified range. The vertical size of the content is dynamically calculated, bu
'* the horizontal is not.
'*
'* @param first_data_cell The initial cell of the range.
'* @param h_size The horizontal size of the range.
'*
Public Function ClearRangeContentHSize(ByVal first_data_cell As Range, ByVal h_size As Long, ByRef clear_flags As ClearContentFlags) As String
    Dim no_of_rows As Long
    
    ' Get the position of the last vertical cell
    no_of_rows = first_data_cell.End(xlDown).Row - first_data_cell.Row + 1
    
    ClearRangeContentHSize = ClearRangeContentSize(first_data_cell, h_size, no_of_rows, clear_flags)
End Function

'**
'* Clear the contents of the specified range. The vertical size of the content is dynamically calculated, bu
'* the horizontal is not.
'*
'* @param first_data_cell The initial cell of the range.
'* @param h_size The horizontal size of the range.
'*
Public Function ClearRangeContentVSize(ByVal first_data_cell As Range, ByVal v_size As Long, ByRef clear_flags As ClearContentFlags) As String
    Dim no_of_columns As Long
    
    ' Get the position of the last horizontal.
    no_of_columns = first_data_cell.End(xlToRight).Column - first_data_cell.Column + 1
    
    ClearRangeContentVSize = ClearRangeContentSize(first_data_cell, no_of_columns, v_size, clear_flags)
End Function

'**
'* Clear the contents of the specified range. The vertical size of the content is dynamically calculated, bu
'* the horizontal is not.
'*
'* @param first_data_cell The initial cell of the range.
'* @param h_size The horizontal size of the range.
'*
Public Function ClearRangeContentSize(ByVal first_data_cell As Range, ByVal h_size As Long, ByVal v_size As Long, ByRef clear_flags As ClearContentFlags) As String
    Dim no_of_rows As Long
    Dim max_row_number As Long
    Dim first_data_cell_row As Long
    Dim first_data_cell_column As Long
    Dim range_to_clear As Range
   
    ' adjust the sizes to avoid limits overbound
    first_data_cell_row = first_data_cell.Row
    If v_size + first_data_cell_row > kMaxRows Then
        v_size = kMaxRows - first_data_cell_row
    End If
    
    first_data_cell_column = first_data_cell.Column
    If h_size + first_data_cell_column > kMaxColumns Then
        h_size = kMaxColumns - first_data_cell_column
    End If
    
    Set range_to_clear = first_data_cell.Resize(v_size, h_size)
    
    On Error Resume Next
    range_to_clear.value = Empty
    range_to_clear.ClearContents
    On Error GoTo 0
    
    If clear_flags.ClearConditionalFormats Then
        range_to_clear.FormatConditions.Delete
    End If
    
    If clear_flags.ClearFormats Then
        'range_to_clear.
    End If
    
    ClearRangeContentSize = range_to_clear.Address
End Function

'**
'* Ensure that the pseudo constants variables are set.
'*
'* The SetConstants function implemented by this library on set the value of
'* the kConstIsSet variable. Users of this library must add the others constants
'* to the body of the SetConstants function.
'*
Public Sub EnsureConstants()
    If kConstIsSet = Empty Then
        SetConstants
    End If
End Sub

'**
'* Set the value of pseudo constants variables.
'*
'Public Sub SetConstants()
'    kConstIsSet = True
    
    ' Project pseudo constants variables must be added above:
    '
    ' Do something like that:
    ' kMyConstant = GetConfiguredValue("kMyConstant")
'End Sub

'**
'* Scans the rows for a columns searching for the specified value and returns the value in the same row based
'* on the specified offset.
'* <p>The scan process stops either if a empty cell is found or if the value is found.
'*
'* @param search_range The first cell to starts the scan. When this method returns contains contains the range
'*                     where the searched value was found and when false is returned contain the last range
'*                     that was searched.
'* @param search_value The value to search for.
'*
'* @return True if the specified value was found; otherwise, false.
'*
Public Function SearchFor(ByRef search_range As Range, ByVal search_value As String) As Boolean
    Dim result As Boolean
    
    On Error GoTo Finally
    
    Do While (search_range.value <> Empty And search_range.Row < kMaxRows)
        If search_range.value = search_value Then
            result = True
            Exit Do
        End If
        
        Set search_range = search_range.Offset(1, 0)
    Loop
    
    GoTo Finally
    
Catch:
    result = False
    
Finally:
    SearchFor = result
End Function

Public Function SetConditionalFormatSize(ByVal first_data_cell As Range, ByVal no_of_columns As Long, ByVal no_of_rows As Long, ParamArray formats_conditions() As Variant) As String
    Dim first_data_cell_row As Long
    Dim range_to_set As Range
    Dim format_conditions_count As Long
    Dim format_condition_pos As Long
    Dim i As Long
   
    ' adjust the sizes to avoid limits overbound
    first_data_cell_row = first_data_cell.Row
    If no_of_rows > kMaxRows + first_data_cell_row Then
        no_of_rows = kMaxRows - first_data_cell_row
    End If
    
    format_conditions_count = UBound(formats_conditions)
    If format_conditions_count > 0 And format_conditions_count < 4 Then
        Set range_to_set = first_data_cell.Resize(no_of_rows, no_of_columns)
        range_to_set.FormatConditions.Delete
        
        format_conditions_count = UBound(formats_conditions)
        For format_condition_pos = 1 To format_conditions_count
            range_to_set.FormatConditions(i) = formats_conditions(i)
        Next
    End If
    
    SetConditionalFormatSize = range_to_set.Address
End Function

Public Function RangeAddr(ByVal data_range As Range)
    RangeAddr = data_range.Address
End Function
