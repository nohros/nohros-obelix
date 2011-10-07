Attribute VB_Name = "obelix_net"
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
' This module contains methods to handle with the network.
'
Option Explicit
Option Base 1

'**
'* Download a binary from a URL and store them into the download folder of the obelix application data folder.
'* <p>
'* The obelix downlaod application data folder is located on %appdata%\nohros\obelix\download
'*
'* @param resource_url The URL of the resource to get.
'* @param The name of the resulting file on disk. ifthe file already exists it will be overwritten.
Public Function GetBinaryFromWeb(ByVal resource_url As String, ByVal file_name As String) As Boolean
    Dim xml_http_request As Object ' XmlHttpRequest Object. Not included to minimize dependencies
    Dim binary_resource() As Byte
    
    GetBinaryFromWeb = False
    
    On Error GoTo Catch
    
    Set xml_http_request = CreateObject("Microsoft.XMLHTTP")
    
    ' open the synchrounous socket to get the data.
    xml_http_request.Open "GET", resource_url, False
    xml_http_request.Send
    
    ' waiting the request to finish
    Do While xml_http_request.ReadyState <> 4
        DoEvents
    Loop
    
    If xml_http_request.Status = 200 Then
        binary_resource = xml_http_request.responseBody
        obelix_io.SaveBinary GetDownloadPathFor(file_name), binary_resource
        
        GetBinaryFromWeb = True
    Else
        LogError xml_http_request.StatusText
        GetBinaryFromWeb = False
    End If
    
    GoTo Finally

Catch:
    LogError Err.Description
    
Finally:
    Set xml_http_request = Nothing
End Function
