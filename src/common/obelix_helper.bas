Attribute VB_Name = "obelix_helper"
' Copyright (c) 2010 Nohros Systems Inc.
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
' This module contains helper methods that performs general tasks.
'
Option Explicit
Option Base 1

'**
'* Call the SQLQueryToArray function and pass the resulting array to the SQLiteTableFromArray method.
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

'**
'* Checks for the existence of resources into the obelix binary folder and download it from
'* the specified URI if it not exists.
'* <p> The name of the resource inside the URI must be specified using the string "$1"(without quotes)
'* The string $1 will be replaced by the name of the resource, when a download is required.
'*
'* @param resources_uri_mask A URI that points to the location where the resources can be downloaded.
'* @param resources_names The name of the resources the check.
Public Function CheckBinaryResources(ByVal resources_uri_mask, ParamArray resources_names() As Variant)
    Dim iterator_i As Integer
    Dim resources_names_size As Long
    Dim resource_name As String
    Dim resource_path As String
    Dim result As Boolean
    
    On Error GoTo Catch
    
    result = True
    
    If IsEmpty(resources_names) Then
        GoTo Finally
    End If
    
    resources_names_size = UBound(resources_names)
    For iterator_i = 0 To resources_names_size ' ParamArray is always zero-based
        resource_name = resources_names(iterator_i)
        If Not obelix_io.ExistsInBin(resource_name) Then
            result = obelix_net.GetBinaryFromWeb(FORMATARTEXTO(resources_uri_mask, resource_name & ".zip"), resources_names(iterator_i))
            
            ' we need to move the downloaded resource to the binary folder
            ' removing the .zip extension
            Name GetDownloadPathFor(resource_name) As GetBinPathFor(resource_name)
        End If
    Next iterator_i
    
    GoTo Finally
    
Catch:
    ' log the error and propagate to the caller
    LogError "[obelix_helper   CheckBinaryResources]   " & Err.Description
    
    Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
Finally:
    CheckBinaryResources = result
End Function

'**
'* Register the specified resource component againts the Windows Registry.
'* <p>The component must be registrable COM Server and must be located into the obelix binary folder.
'* The regsvrex app will be used to do the registration process and it must be location into the obelix
'* folder too.
'*
'* @param resource_name The name of the resource to register.
'*
'* @return true if the registration process succeeds; otherwise false.
Public Function RegisterBinary(ByVal resource_name As String) As Boolean
    Dim resource_path As String
    Dim regsvrex_path As String
    Dim result As Double
    
    On Error GoTo Catch
    
    resource_path = obelix_io.GetBinPathFor(resource_name)
    regsvrex_path = obelix_io.GetBinPathFor(kRegSvrExFileName)
    
    ' register the resource
    result = Shell(FORMATARTEXTO("""$1"" /c ""$2""", regsvrex_path, resource_path))
    
    RegisterBinary = True
    
    GoTo Finally
    
Catch:
    LogError "[obelix_helper   RegisterBinary]   " & Err.Description
    RegisterBinary = False
    
Finally:
End Function

Public Function VBADateToUnixSeconds(ByVal vba_date As Date)
    VBADateToUnixSeconds = DateDiff("s", #1/1/1970#, vba_date)
End Function

Public Function UnixSecondsToVBADate(ByVal unix_seconds As Long)
    UnixSecondsToVBADate = DateAdd("s", unix_seconds, #1/1/1970#)
End Function

Public Function LngNotIn(ByVal lng_value As Long, ByRef lng_arr() As Long) As Boolean
    Dim no_of_elements As Long
    Dim iterator_i As Integer
    Dim result As Boolean
    
    result = True
    
    no_of_elements = UBound(lng_arr)
    
    For iterator_i = 1 To no_of_elements
        If lng_value = lng_arr(iterator_i) Then
            result = False
        End If
    Next iterator_i

    LngNotIn = result
End Function

Public Function LngIn(ByVal lng_value As Long, ByRef lng_arr() As Long) As Boolean
    Dim no_of_elements As Long
    Dim iterator_i As Integer
    Dim result As Boolean
    
    result = False
    
    no_of_elements = UBound(lng_arr)
    
    For iterator_i = 1 To no_of_elements
        If lng_value = lng_arr(iterator_i) Then
            result = True
            Exit For
        End If
    Next iterator_i

    LngIn = result
End Function

Public Function TryParseDate(ByVal expr As String, ByVal date_format As String, ByRef parsed_date As Date) As Boolean
    Dim iterator_i As Long
    Dim iterator_j As Long
    Dim iterator_k As Long
    
    Dim date_format_length As Integer
    Dim current_date_format_char As String
    Dim last_date_format_char As String
    Dim current_date_format_pos As Long
    Dim date_part_pos As Long
    Dim current_expr_pos As Long
    
    Dim current_token As String
    Dim next_token As String
    Dim token_dynamic_arr As String
    
    Dim tokens() As String
    Dim valid_tokens() As String
    Dim date_parts(3) As Long
    
    Dim no_of_valid_tokens As Long
    Dim no_of_tokens As Long
    Dim tokens_size As Long
    Dim current_date_part_size As Long
    Dim last_expr_pos As Long
    
    Dim result As Boolean
    Dim index_of As Long
    Dim is_valid_token As Boolean
    
    On Error GoTo Catch
    
    Const kYearPos As Integer = 1
    Const kMonthPos As Integer = 2
    Const kDayPos As Integer = 3
    
    no_of_valid_tokens = 6
    current_expr_pos = 1
    last_expr_pos = 1
    
    ReDim valid_tokens(no_of_valid_tokens)
    
    valid_tokens(1) = "d"
    valid_tokens(2) = "dd"
    valid_tokens(3) = "m"
    valid_tokens(4) = "mm"
    valid_tokens(5) = "yy"
    valid_tokens(6) = "yyyy"
    
    date_format_length = Len(date_format)
    current_date_format_pos = 1
    last_date_format_char = ""
    current_date_format_char = ""
    
    ' parse the format
    For iterator_i = 1 To date_format_length
        last_date_format_char = current_date_format_char
        current_date_format_char = Mid(date_format, current_date_format_pos, 1)
        
        If last_date_format_char = current_date_format_char Or last_date_format_char = "" Then
            current_token = current_token & current_date_format_char
        Else
            ' put the current_token into the array and starts capturing a new one
            tokens_size = tokens_size + Len(current_token)
            token_dynamic_arr = token_dynamic_arr & current_token & "&"
            current_token = current_date_format_char
        End If
        current_date_format_pos = current_date_format_pos + 1
    Next iterator_i
    
    ' validate the date against the format
    tokens = Split(token_dynamic_arr, "&")
    no_of_tokens = UBound(tokens) + 1
    
    ' store the last token into the dynamic array
    tokens(no_of_tokens - 1) = current_token
    tokens_size = tokens_size + Len(current_token)
    
    If Len(expr) < tokens_size Then
        result = False
        GoTo Finally
    End If
    
    For iterator_i = 0 To no_of_tokens - 1
        current_token = tokens(iterator_i)
        If iterator_i < no_of_tokens - 1 Then
            next_token = tokens(iterator_i + 1)
        Else
            next_token = ""
        End If
        
        ' check if the current token is a datepart
        For iterator_j = 1 To no_of_valid_tokens
            If current_token = valid_tokens(iterator_j) Then
                ' get the position to store the current datepart
                Select Case current_token
                    Case "d", "dd"
                        date_part_pos = kDayPos
                        
                    Case "m", "mm"
                        date_part_pos = kMonthPos
                        
                    Case "yy", "yyyy"
                        date_part_pos = kYearPos
                End Select
                
                is_valid_token = False
                For iterator_k = 1 To no_of_valid_tokens
                    If next_token = valid_tokens(iterator_k) Then
                        is_valid_token = True
                        Exit For
                    End If
                Next iterator_k
                
                If is_valid_token Then
                    current_date_part_size = Len(current_token)
                    last_expr_pos = current_expr_pos + current_date_part_size
                    current_expr_pos = current_expr_pos + current_date_part_size
                Else
                    If iterator_i = no_of_tokens - 1 Then
                        index_of = Len(expr) + 1
                    Else
                        index_of = InStr(current_expr_pos, expr, next_token)
                    End If

                    current_date_part_size = index_of - current_expr_pos
                    last_expr_pos = current_expr_pos
                    current_expr_pos = index_of + Len(current_token)
                End If
                date_parts(date_part_pos) = CLng(Mid(expr, last_expr_pos, current_date_part_size))

                GoTo LBL_continue_i
            End If
        Next iterator_j
LBL_continue_i:
    Next iterator_i
    
    parsed_date = DateSerial(date_parts(kYearPos), date_parts(kMonthPos), date_parts(kDayPos))
    
    result = True
    
    GoTo Finally
    
Catch:
    result = False

Finally:
    TryParseDate = result
End Function
