VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortedDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private sorted_dictionary_ As Object 'dhRichClient3.cSortedDictionary'

Private Sub Class_Initialize()
    obelix_win.CreateRegFreeObject _
        obelix_consts.kCollectionClassName, _
        obelix_io.GetBinPathFor( _
            obelix_consts.kdhRichClientDLLName _
        ), _
        sorted_dictionary_
End Sub

Public Sub Add(ByVal item_key As String, ByVal item_value As Variant)
    sorted_dictionary_.Add item_value, item_key
End Sub

Public Sub Clear()
    sorted_dictionary_.RemoveAll
End Sub

Public Property Get Item(ByVal key As String, ByVal default As Variant) As String
    If sorted_dictionary_.Exists(key) Then
        Item = sorted_dictionary_.Item(key)
    Else
        Item = default
    End If
End Property

Public Property Get ItemByIndex(ByVal index As Long) As String
    ItemByIndex = sorted_dictionary_.ItemByIndex(index)
End Property

Public Property Get KeyByIndex(ByVal index As Long) As String
    KeyByIndex = sorted_dictionary_.KeyByIndex(index)
End Property

Public Property Get Count() As Long
    Count = sorted_dictionary_.Count
End Property

Public Property Get Keys() As String()
    Dim iterator_row As Long
    Dim keys_arr() As String
    Dim keys_count As Long
    
    keys_count = sorted_dictionary_.Count
    
    If keys_count > 0 Then
        ReDim keys_arr(keys_count)
    
        For iterator_row = 1 To keys_count
            keys_arr(iterator_row) = sorted_dictionary_.KeyByIndex(iterator_row - 1)
        Next iterator_row
    Else
        ReDim keys_arr(1)
    End If
    
    Keys = keys_arr
End Property


