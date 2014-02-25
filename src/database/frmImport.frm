VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "SQL Import Properties"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   OleObjectBlob   =   "frmImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const kMinTop As Integer = 18
Const kControlHeight As Integer = 12

Private destionations() As combobox_events
Private sources() As label_events
Private masks() As label_events

Private Sub CommandButton2_Click()
    frmImport.Hide
End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Activate()
    Dim sql_connection As Connection
    Dim rows() As Variant
    Dim fields() As String
    Dim field_types() As String
    Dim selection As Range
    Dim top As Integer
    
    On Error GoTo Catch
    
    Set sql_connection = obelix_mssql.OpenConnection(obelix.connection_string_)
    
    If obelix_mssql.SQLQueryToArray( _
        "select name from sys.tables", _
        sql_connection, rows, fields, field_types) Then
        tables.List = Application.Transpose(rows)
    End If
    
    Set selection = Application.selection
    
    top = kMinTop
    For i = 1 To selection.Columns.Count
        AddControls selection.rows(1).Columns(i), top
        top = top + kControlHeight - 1
    Next i
    
    GoTo Finally
    
Catch:
    MsgBox Err.Description
Finally:
End Sub

Private Function AddControls(ByVal header As Range, ByVal top As Integer, ByVal position As Integer)
    Dim source As MSForms.Label
    Dim destination As MSForms.ComboBox
    Dim mask As MSForms.Label
    
    Set source = transformations.Controls.Add("Forms.Label.1")
    With source
        .Caption = " " + header.Value
        .top = top
        .Left = 0
        .Width = 156
        .BackColor = RGB(255, 255, 255)
        .Height = kControlHeight
        .BorderColor = RGB(191, 191, 191)
        .BorderStyle = fmBorderStyleSingle
    End With
    
    Set destination = transformations.Controls.Add("Forms.Combobox.1")
    With destination
        .top = top
        .Left = 155
        .Height = kControlHeight
        .Width = 161
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = RGB(191, 191, 191)
    End With
    
    'Set mask = transformations.Controls.Add("Forms.Label.1")
    'With mask
        '.top = top
        '.Left = destination.Left
        '.Height = kControlHeight
        '.Width = destination.Width
        '.BackColor = source.BackColor
        '.BorderStyle = source.BorderStyle
        '.BorderColor = source.BorderColor
    'End With
End Function
