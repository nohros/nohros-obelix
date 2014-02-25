VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConnection 
   Caption         =   " String de Conexão"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   OleObjectBlob   =   "frmConnection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private server_ As String
Private login_ As String
Private password_ As String
Private initial_catalog_ As String

Private Sub btnCancel_Click()
    frmConnection.Hide
End Sub

Private Sub btnOK_Click()
    If Len(server_) = 0 Or _
        Len(login_) = 0 Or _
        Len(password_) = 0 Or _
        Len(initial_catalog_) = 0 Then
        MsgBox "Preencha todos os campos"
    Else
        obelix.connection_string_ = _
            "Data Source=" & server_ & _
            ";Initial Catalog=" & initial_catalog_ & _
            ";User ID=" & login_ & _
            ";Password=" & password_ & ";"
            frmConnection.Hide
    End If
End Sub

Private Sub database_Change()
    initial_catalog_ = database.Text
End Sub

Private Sub txtLoginID_Change()
    login_ = txtLoginID.Text
    txtPassword.Text = Empty
    database.Clear
End Sub

Private Sub txtPassword_Change()
    password_ = txtPassword.Text
    database.Clear
End Sub

Private Sub txtServer_Change()
    server_ = txtServer.Text
    database.Clear
End Sub

Private Sub LoadDatabases()
    
    On Error GoTo Catch
    
    Dim connection_string As String
    Dim sql_connection As Connection
    Dim rows() As Variant
    Dim fields() As String
    Dim fields_types() As String
    
    If Len(server_) = 0 Or Len(login_) = 0 Or Len(password_) = 0 Then
        GoTo Finally
    End If
    
    database.Enabled = False

    connection_string = _
        "Data Source=" & server_ & _
        ";User ID=" & login_ & _
        ";Password=" & password_ & ";"
            
    Set sql_connection = obelix_mssql.OpenConnection(connection_string)
    If obelix_mssql.SQLQueryToArray( _
        "select name from sys.databases", _
        sql_connection, rows, fields, fields_types) Then
        database.List = Application.Transpose(rows)
    End If
    
Catch:
    'LogError Err
    
Finally:
    obelix_mssql.CloseConnection sql_connection
    database.Enabled = True
End Sub

Private Sub txtServer_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    LoadDatabases
End Sub

Private Sub txtLoginID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    LoadDatabases
End Sub

Private Sub txtPassword_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    LoadDatabases
End Sub


Private Sub UserForm_Activate()
    txtServer.Text = server_
    txtLoginID.Text = login_
    txtPassword.Text = Empty
    database.Text = initial_catalog_
    
    If txtServer.TextLength = 0 Then
        txtServer.SetFocus
    ElseIf txtLoginID.TextLength = 0 Then
        txtLoginID.SetFocus
    ElseIf txtPassword.TextLength = 0 Then
        txtPassword.SetFocus
    End If
End Sub
