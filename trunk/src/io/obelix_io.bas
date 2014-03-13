Attribute VB_Name = "obelix_io"
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
' This module contains methods to handle IO related operation
'
Option Explicit
Option Base 1

Public Const kNohrosRootFoldername As String = "nohros"
Public Const kObelixRootFolderName As String = kNohrosRootFoldername + "\obelix"

Public Const kObelixBinFolderName As String = kObelixRootFolderName + "\bin"
Public Const kObelixDownloadFolderName As String = kObelixRootFolderName + "\download"
Public Const kObelixSQLiteFolderName As String = kObelixRootFolderName + "\sqlite"

'**
'* Creates the obelix data folder structure. The obelix data folders
'* will be created into the current user %appdata%. The folder
'* structure follows the pattern above:
'*
'* %appdata%
'*   obelix
'*     download
'*
'* If the folder structre already exists we will do nothing and no errors will be raised.
'* To use ths module you need to add a reference to the Microsoft Scripting Runtime
'* library. Usually this is located at "%systemroot%\system32\scrrun.dll"
Public Function CreateObelixDataFolderStructure() As Boolean
    Dim fs_object As FileSystemObject
    Dim app_data_path As String
    Dim succeeded As Boolean
    Dim no_of_folders_in_structure As Integer
    Dim folders_struct() As String
    Dim iterator_i As Integer
    
    On Error GoTo Catch
    
    ' pessimistic, false until true
    succeeded = False
    
    app_data_path = Environ$("AppData") & "\"
    
    Set fs_object = New FileSystemObject
    
    ' set the number of paths that the structure contains
    no_of_folders_in_structure = 5
    
    ' create a array representing the folder structure
    ' ex.
    '   arr(0)
    '     arr(1)
    '     arr(2)
    '   arr(3)
    '
    ' This array will be readed sequentially from top to down
    ReDim folders_struct(no_of_folders_in_structure)
    
    ' creating the folder's paths  and adding them to the structure tree
    folders_struct(1) = app_data_path & kNohrosRootFoldername
    folders_struct(2) = app_data_path & kObelixRootFolderName
    folders_struct(3) = app_data_path & kObelixBinFolderName
    folders_struct(4) = app_data_path & kObelixDownloadFolderName
    folders_struct(5) = app_data_path & kObelixSQLiteFolderName
    
    For iterator_i = 1 To no_of_folders_in_structure
        succeeded = CreateFolderIfNotExists(fs_object, folders_struct(iterator_i))
        
        ' only create the next folder if the creation of the previous succeeds
        If Not succeeded Then
            Exit For
        End If
    Next iterator_i
    
    GoTo Finally
    
Catch:
    LogError Err
    
Finally:
    CreateObelixDataFolderStructure = succeeded
    
    Set fs_object = Nothing
End Function

'**
'* Checks if a given path exists and create it if not.
'* <p>
'* If a error occurs it will be logged and the error propagated to the caller.
'*
'* @param A FileSystemObject used to manipulate the file system.
'* @param The path to the folder to check.
'*
'* @return true if no erros has been occurred. otherwise false.
Private Function CreateFolderIfNotExists(ByVal fs_object As FileSystemObject, ByVal folder_path As String) As Boolean
    On Error GoTo Catch
    
    CreateFolderIfNotExists = False
    
    If Not fs_object.FolderExists(folder_path) Then
        fs_object.CReateFolder folder_path
    End If
    
    CreateFolderIfNotExists = True
    
    GoTo Finally
    
Catch:
    LogError Err
    
    Err.Raise Err.number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
Finally:
End Function

'**
'* Saves the content of a byte array into disk.
'* <p> The content will be created into the obelix download folder. If the file
'* already exists your content will be overwirtten
'*
Public Function SaveBinary(ByVal file_name As String, ByRef binary_data() As Byte) As Boolean
    Dim file_path As String
    Dim fs_object As FileSystemObject
    Dim file_handle As Long
    
    On Error GoTo Catch
    
    SaveBinary = False
    
    Set fs_object = New FileSystemObject
    If fs_object.FileExists(file_name) Then
        fs_object.DeleteFile file_name
    End If
    
    file_handle = FreeFile
    Open file_name For Binary As #file_handle
    Put #file_handle, , binary_data
    
    SaveBinary = True
    
Catch:
Finally:
    Close #file_handle
End Function

'**
'* Gets the path where a given file could be stored
'* <p>
'* The path returned will be equals to the concatenation of the obelix
'* folder path and the specified file name.
'*
'* @param The name of the file to get the path from.
Public Function GetDownloadPathFor(ByVal file_name As String) As String
    GetDownloadPathFor = GetDownloadPath & "\" & file_name
End Function

'**
'* Gets the path for the obelix download folder
'* <p>
'* The returned path does not contains a trailing slash
Public Function GetDownloadPath() As String
    GetDownloadPath = Environ("AppData") & "\" & kObelixDownloadFolderName
End Function

'**
'* Gets the path for the obelix bin folder
'* <p>
'* The returned path does not contains a trailing slash
Public Function GetBinPath() As String
    GetBinPath = Environ("AppData") & "\" & kObelixBinFolderName
End Function

'**
'* Gets the path where a given file could be stored
'* <p>
'* The path returned will be equals to the concatenation of the obelix
'* folder path and the specified file name.
'*
'* @param The name of the file to get the path from.
Public Function GetBinPathFor(ByVal file_name As String) As String
    GetBinPathFor = GetBinPath & "\" & file_name
End Function

'**
'* Gets the path for the obelix sqlite folder
'* <p>
'* The returned path does not contains a trailing slash
Public Function GetSQLitePath() As String
    GetSQLitePath = Environ("AppData") & "\" & kObelixSQLiteFolderName
End Function

'**
'* Gets the path where a given file could be stored
'* <p>
'* The path returned will be equals to the concatenation of the obelix
'* folder path and the specified file name.
'*
'* @param The name of the file to get the path from.
Public Function GetSQLitePathFor(ByVal file_name As String) As String
    GetSQLitePathFor = GetSQLitePath & "\" & file_name
End Function

'**
'* Check if a given file exists into the obelix bin folder
Public Function ExistsInBin(ByVal file_name As String) As Boolean
    ExistsInBin = Len(Dir(GetBinPathFor(file_name), vbNormal)) > 0
End Function

'**
'* Check if a given file exists into the obelix sqlite folder
Public Function ExistsInSQLite(ByVal file_path As String)
    ExistsInSQLite = Len(Dir(GetSQLitePathFor(file_path), vbNormal)) > 0
End Function

Public Function QuotePath(ByVal path As String)
    QuotePath = Chr(34) & path & Chr(34)
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
            result = obelix_net.GetBinaryFromWeb(FormatarTexto(resources_uri_mask, resource_name & ".zip"), resources_names(iterator_i))
            
            ' we need to move the downloaded resource to the binary folder
            ' removing the .zip extension
            Name GetDownloadPathFor(resource_name) As GetBinPathFor(resource_name)
        End If
    Next iterator_i
    
    GoTo Finally
    
Catch:
    ' log the error and propagate to the caller
    LogError "[obelix_helper   CheckBinaryResources]   " & Err.Description
    
    Err.Raise Err.number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
    
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
    result = Shell(FormatarTexto("""$1"" /c ""$2""", regsvrex_path, resource_path))
    
    RegisterBinary = True
    
    GoTo Finally
    
Catch:
    LogError "[obelix_helper   RegisterBinary]   " & Err.Description
    RegisterBinary = False
    
Finally:
End Function
