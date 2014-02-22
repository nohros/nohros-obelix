Attribute VB_Name = "obelix_consts"
Option Explicit

Public Const kSQLiteEngineFileName As String = "sqlite36_engine.dll"
Public Const kRegSvrExFileName = "regsvrex.exe"
Public Const kSQLiteManifest As String = "dhRichClient3.dll.manifest"
Public Const kSQLiteCOMClassName As String = "dhRichclient3.cConnection"
Public Const kSQLiteConnectionClassName As String = "cConnection"
Public Const kCollectionClassName As String = "cCollection"
Public Const kSortedDictionaryClassName As String = "cSortedDictionary"
Public Const kCryptClassName As String = "cCrypt"
Public Const kdhRichClientDLLName As String = "dhRichClient3.dll"
Public Const kSQLiteFileName As String = kdhRichClientDLLName
Public Const kdhDirectCOMDLLName As String = "DirectCOM.dll"

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
