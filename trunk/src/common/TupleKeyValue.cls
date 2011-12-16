VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TupleKeyValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private sorted_dictionary_ As Object 'dhRichClient3.cSortedDictionary'
Private join_char_ As String
Private key_prefix_ As String

Private Sub Class_Initialize()
    obelix_win.CreateRegFreeObject _
        obelix_consts.kSortedDictionaryClassName, _
        obelix_io.GetBinPathFor( _
            obelix_consts.kdhRichClientDLLName _
        ), _
        sorted_dictionary_
    
    join_char_ = "_"
    key_prefix_ = ""
End Sub

Public Sub Add(ByVal part_a As Long, ByVal part_b As Long, ByVal value As String)
    sorted_dictionary_.Add key(part_a, part_b), value
End Sub

Public Sub Clear()
    sorted_dictionary_.RemoveAll
End Sub
Public Function key(ByVal part_a As Long, ByVal part_b As Long) As String
    key = key_prefix_ & CStr(part_a) & join_char_ & CStr(part_b)
End Function

Public Property Get Item(ByVal part_a As Long, ByVal part_b As Long) As String
    Dim str_key As String
    
    str_key = key(part_a, part_b)
    If sorted_dictionary_.Exists(str_key) Then
        Item = sorted_dictionary_.Item(str_key)
    Else
        Item = ""
    End If
End Property

Public Property Get Keys() As String()
    Dim iterator_row As Long
    Dim keys_arr() As String
    Dim keys_count As Long
    
    keys_count = sorted_dictionary_.Count
    For iterator_row = 0 To keys_count
        keys_arr(iterator_row + 1) = sorted_dictionary_.ItemByIndex(iterator_row)
    Next iterator_row
    
    Keys = keys_arr
End Property

Public Property Get JoinChar() As String
    JoinChar = join_char_
End Property

Public Property Let JoinChar(ByVal join_char As String)
    join_char_ = join_char
End Property

Public Property Get KeyPrefix() As String
    KeyPrefix = key_prefix_
End Property

Public Property Let KeyPrefix(ByVal key_prefix As String)
    key_prefix_ = key_prefix
End Property
