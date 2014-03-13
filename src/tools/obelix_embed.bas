Attribute VB_Name = "obelix_embed"
' Copyright (c) 2010 by Nohros Systems Inc.
' Copyright (c) 2003 by Dermot Balson.
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
' To restore embedded data, you need to include the MustEmbed module in the file you have
' embedded data in, and, if you have used compression, you will also need the MustCompress
' module.
'
' To delete embedded data, simply delete the sheet containing the cellnotes (make it visible
' first).
'
' This code requires no third party software. It uses one Windows API call.
' It is free to use as long as you don't claim credit for it yourself.
'
' This code was written by Dermot Balson and revised by the Nohros Developers.
'
Option Explicit
Option Base 1

Function EmbedData(container As Worksheet, file_to_embed As String, ByVal column_index As Long, Optional compress As Boolean, Optional check As Boolean) As Boolean
    Dim B1() As Byte, B2() As Byte
    Dim N As Long, u As Long
    Dim nBytes As Long
    Dim data As String
    Dim checked_data() As Byte
    
    'get datafile & compress it if required
    If compress Then
      nBytes = ReadFile(file_to_embed, B1())
      'compress it if required
      CompressData B1(), B2()
    Else
      nBytes = ReadFile(file_to_embed, B2())
    End If
    
    'remap characters that won't go into cell comments
    'this adds about 12% to the text length
    Remap B2(), B1()
    data = StrConv(B1(), vbUnicode)
    Erase B1()
    Erase B2()
    
    'embed data in chunks of 32767 characters
    u = 1
    N = 1
    
    ' write the name of the embed file
    'container.Cells(1, column_index).Value = GetFileName(file_to_embed)
    
    Do
      If u > Len(data) Then Exit Do
      N = N + 1
      container.Cells(N, column_index).AddComment Mid$(data, u, 32767)
      u = u + 32767
    Loop
    
    'hide data sheet
    container.Visible = xlVeryHidden
    'container.Parent
    
    'check it worked, if requested
    If check Then
      checked_data = RecoverData(container, compress, column_index)
      
      'return success flag
      nBytes = ReadFile(file_to_embed, B1())
      data = StrConv(B1(), vbUnicode)
      EmbedData = (data = StrConv(checked_data(), vbUnicode))
    Else
      EmbedData = True 'if not testing, just return success flag
    End If
End Function

Function RecoverData(ByVal container As Worksheet, ByVal is_compressed As Boolean, ByVal column_index As Long) As Byte()
    Dim B1() As Byte, B2() As Byte
    Dim data As String
    Dim chunk As Comment
    Dim data_row As Range
    Dim i As Long
    Dim j As Long
    
    container.Visible = True
    For i = 2 To 65536
        Set data_row = container.Cells(i, column_index)
        If data_row.Comment Is Nothing Then
            Exit For
        End If
        data = data & data_row.Comment.Text
    Next i
    
    'convert to byte array and remap
    ReDim B1(Len(data))
    CopyMemory B1(1), ByVal data, Len(data)
    Demap B1(), B2()
    
    'decompress if necessary
    If is_compressed = True Then
      DecompressData B2(), B1()
      RecoverData = B1()
    Else
      RecoverData = B2()
    End If
    
    container.Visible = xlSheetVeryHidden
    
    Erase B1(), B2()
End Function

Sub Remap(inB() As Byte, outB() As Byte)
    Dim u&, m&, N&
    Dim a(0 To 255) As Long
    
    a(0) = 1
    a(1) = 1
    a(128) = 1
    a(130) = 1
    a(131) = 1
    a(132) = 1
    a(133) = 1
    a(134) = 1
    a(135) = 1
    a(136) = 1
    a(137) = 1
    a(138) = 1
    a(139) = 1
    a(140) = 1
    a(142) = 1
    a(145) = 1
    a(146) = 1
    a(147) = 1
    a(148) = 1
    a(149) = 1
    a(150) = 1
    a(151) = 1
    a(152) = 1
    a(153) = 1
    a(154) = 1
    a(155) = 1
    a(156) = 1
    a(157) = 1
    a(158) = 1
    a(159) = 1
    
    u = 0
    m = UBound(inB)
    ReDim outB(m * 1.2) As Byte
    
    For N = 1 To m
      u = u + 1
      If a(inB(N)) > 0 Then
        outB(u) = 1
        u = u + 1
        outB(u) = inB(N) + 40
      Else
        outB(u) = inB(N)
      End If
    Next N
    
    ReDim Preserve outB(u)
End Sub

Sub Demap(inB() As Byte, outB() As Byte)
    Dim u&, m&, N&

    u = 0
    m = UBound(inB)
    ReDim outB(m) As Byte
    For N = 1 To m
      u = u + 1
      If inB(N) = 1 Then
        N = N + 1
        outB(u) = inB(N) - 40
      Else
        outB(u) = inB(N)
      End If
    Next N
    ReDim Preserve outB(u)
End Sub

Function WriteEmbedToDisk(ByVal container As Workbook, ByVal name As String, ByVal BASE_PATH As String)
    Dim embed_sheet As Worksheet
    Dim data_column As Range
    Dim i As Long
    Dim data() As Byte
    Dim data_name As String
    
    Set embed_sheet = container.Sheets("MUSTOLE")
    
    For Each data_column In embed_sheet.Columns
        data_name = StrConv(data_column.Cells(1, 1).value, vbLowerCase)
        If data_name = Empty Then _
            Exit For
        
        If StrConv(name, vbLowerCase) = data_name Or IsEmpty(name) Then
            data = RecoverData(embed_sheet, True, data_column.Column)
            WriteToDisk data_name, data, BASE_PATH
        End If
    Next data_column
End Function

Public Function WriteToDisk(ByVal file_name As String, ByRef data() As Byte, ByVal BASE_PATH)
    Dim fs As FileSystemObject
    Dim file_path As String
    Dim file_stream As TextStream
   
    Set fs = New FileSystemObject
    file_path = fs.BuildPath(BASE_PATH, file_name)
    
    ' Create the file on the disk
    Set file_stream = fs.CreateTextFile(file_path, True)
        
    ' Convert the binary data to string and write them to the file
    file_stream.Write StrConv(data, vbUnicode)
        
    file_stream.Close
End Function
