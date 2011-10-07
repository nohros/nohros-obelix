Attribute VB_Name = "obelix_sql"
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
' This module contains methods that is used to encode/decode data to be used by ad-hoc SQL
' queries

Option Base 1
Option Explicit

'**
'* Encodes a string replacing some unwanted characters.
'* <p> This method removes the following characters form the specified text:
'* <p> ['](singe quote)
'*
'* @param The text to escape
Public Function Escape(ByVal Text As String) As String
    Escape = Replace(Text, "'", "")
End Function

'**
'* Encodes a string replacing some unwanted characters and add a comma to the end of the string.
'* <p> This method removes the following characters form the specified text:
'* <p> ['](singe quote)
'* @param The text to escape
Public Function EscapeWithTralingComma(ByVal Text As String) As String
    EscapeWithTralingComma = Escape(Text) & ","
End Function

'**
'* Removes the last character from a string. It is usually used to remove the last
'* comma from a SQL command when this command is build inside a loop.
'*
'* @param text The string to remove the last character.
Public Function RemoveTrailing(ByVal Text As String)
    RemoveTrailing = Mid(Text, 1, Len(Text) - 1)
End Function

'**
'* Adds a single quote to the begining and the end of the specified string.
'*
'* @param The string to add the single quotes.
Public Function Quote(ByVal Text As String)
    Quote = "'" & Text & "'"
End Function

'**
'* Adds a single quote to the begining of the specified string and a single quote and comma to the end of the string.
'*
'* @param The string to add the single quotes.
Public Function QuoteWithTralingComma(ByVal Text As String)
    QuoteWithTralingComma = "'" & Text & "',"
End Function

'**
'* Adds a single quote to the begining of the specified string and a single quote and comma to the end of the string.
'*
'* @param The string to add the single quotes.
Public Function DataTypeToSQLiteType(ByVal data_type_enum As Integer) As String
    Select Case data_type_enum
        Case 20, 8, 5, 3, 131 'adBigInt, adBSTR, adDouble, adInteger
            DataTypeToSQLiteType = "integer"
        Case Else
            DataTypeToSQLiteType = "text"
    End Select
End Function

Public Function SQLArrayToSimpleArray(ByRef sql_array() As Variant, Optional ByVal column_offset As Long = 0) As Variant
    Dim no_of_rows As Long
    Dim iterator_i As Long
    Dim simple_array() As Variant
    
    no_of_rows = UBound(sql_array, 2) + 1
    
    ReDim simple_array(no_of_rows)
    
    For iterator_i = 1 To no_of_rows
        simple_array(iterator_i) = sql_array(column_offset, iterator_i - 1)
    Next iterator_i
    
    SQLArrayToSimpleArray = simple_array
End Function

Public Function SQLArrayToSimpleStringArray(ByRef sql_array() As Variant, Optional ByVal column_offset As Long = 0) As Variant
    Dim no_of_rows As Long
    Dim iterator_i As Long
    Dim simple_array() As String
    
    no_of_rows = UBound(sql_array, 2) + 1
    
    ReDim simple_array(no_of_rows)
    
    For iterator_i = 1 To no_of_rows
        simple_array(iterator_i) = CStr(sql_array(column_offset, iterator_i - 1))
    Next iterator_i
    
    SQLArrayToSimpleStringArray = simple_array
End Function
