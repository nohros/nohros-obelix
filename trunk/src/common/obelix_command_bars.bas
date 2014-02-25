Attribute VB_Name = "obelix_command_bars"
Private Const kSQLCommandBar As String = "SQL Bar"
Private Const kSQLImportButton As String = "SQL Import"
Private Const kCellCommandBar As String = "Cell"

Public Function CreateCellContextMenu()
    Dim context_menu As CommandBar
    Dim context_menu_controls As CommandBarControls
    Dim sql_import_menu As CommandBarControl
    
    Call DeleteCellContextMenu
    
    Set context_menu = Application.CommandBars(kCellCommandBar)
    Set context_menu_controls = context_menu.Controls
    
    With context_menu_controls.Add(Type:=msoControlButton, before:=context_menu_controls.Count + 1)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "SQLImport"
        .FaceId = 0
        .Caption = "Import to SQL..."
        .Tag = kSQLCommandBar
    End With
End Function

Public Function DeleteCellContextMenu()
   Dim context_menu As CommandBar
   Dim control As CommandBarControl
   
   Set context_menu = Application.CommandBars(kCellCommandBar)
   For Each control In context_menu.Controls
    If control.Tag = kSQLCommandBar Then
        control.Delete
    End If
   Next control
End Function

Private Function SQLImport()
    obelix.connection_string_ = Empty
    frmConnection.Show
    
    If Not connection_string_ = Empty Then
        frmImport.Show
    End If
End Function

