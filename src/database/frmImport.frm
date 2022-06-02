VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImport 
   Caption         =   "SQL Import Properties"
   ClientHeight    =   5880
   ClientLeft      =   50
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
Option Explicit

Const kMinTop As Integer = 18
Const kControlHeight As Integer = 12
Const kMapGridName = "mapGrid"

Private grid_ As vbalGrid
Private selection_ As Range

Private Sub CommandButton2_Click()
    frmImport.Hide
End Sub

Private Sub tables_Change()
  Dim sql_connection As Connection
  Dim rows() As Variant
  Dim fields() As String
  Dim field_types() As String
  
  On Error GoTo catch
  
  Set sql_connection = obelix_mssql.OpenConnection(obelix.connection_string_)
  If obelix_mssql.SQLQueryToArray( _
    "select name from sys.columns where object_name(object_id) = '" & tables.value & "'", _
    sql_connection, rows, fields, field_types) Then
  End If
  
  Map grid_, selection_, fields
  
  GoTo finally
  
catch:
  LogError Err
  ReportError Err
  
finally:
  Set sql_connection = Nothing
End Sub

Private Sub UserForm_Activate()
  Dim sql_connection As Connection
  Dim rows() As Variant
  Dim fields() As String
  Dim field_types() As String
  Dim top As Integer
  Dim grid As Object
    
  On Error GoTo catch
      
  Set selection_ = Application.selection
  Set grid_ = CreateGrid
    
  Set sql_connection = obelix_mssql.OpenConnection(obelix.connection_string_)
  If obelix_mssql.SQLQueryToArray( _
      "select name from sys.tables", _
      sql_connection, rows, fields, field_types) Then
      tables.List = Application.Transpose(rows)
  End If
  
  GoTo finally
    
catch:
  LogError Err

finally:
  Set sql_connection = Nothing
  
  obelix_logger.ThrowIfNeeded Err
End Sub

Private Sub Map(ByVal grid As Object, ByVal selection As Range, ByRef fields() As String)
  Dim i As Integer
  Dim selected_columns() As Variant
  Dim column_name As String
  Dim row As Object
  Dim rows As Long
  
  On Error GoTo catch
  
  selected_columns = selection.Resize(1).value
  rows = UBound(selected_columns, 2)
  
  grid.rows = rows
  grid.Redraw = False

  For i = 1 To rows
    column_name = Trim(selected_columns(1, i))
    If column_name = Empty Then
      Err.Raise vbError + 500, "frmImport.Map", "The name of the column [" & CStr(i) & "] is invalid."
    End If
    
    grid_.CellDetails i, 1, column_name, , , , RGB(10, 15, 16)
  Next i
  
  grid.Redraw = True
  
catch:
finally:
  obelix_logger.ThrowIfNeeded Err
End Sub
Private Function CreateGrid() As Object
  On Error GoTo catch
  
  Set CreateGrid = frmImport.controls.Add("vbAcceleratorGrid6.vbalGrid", kMapGridName, True)
  With CreateGrid
    .Redraw = False
    
    .top = 90
    .Width = 318
    .Height = 170
    .Left = 12
    
    .Editable = True
    .Gridlines = True
    .DefaultRowHeight = 24
    
    .AddColumn vKey:="Source", sHeader:="Source", lColumnWidth:=210
    .AddColumn vKey:="Destination", sHeader:="Destination", lColumnWidth:=210
    
    .SetHeaders
    
    .Redraw = True
  End With
  
  GoTo finally

catch:
finally:
End Function

Private Sub UserForm_Deactivate()
  Clear
End Sub

Private Sub Clear()
  Set grid_ = Nothing
  Set selection_ = Nothing
End Sub

Private Sub UserForm_Terminate()
  Clear
End Sub
