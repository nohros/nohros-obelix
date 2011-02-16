Attribute VB_Name = "MustWnd"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal Indice As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                                       ByVal hwnd As Long, _
                                       ByVal nIndex As Long, _
                                       ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" ( _
                                     ByVal hwnd As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
                                                   ByVal lpWindowName As String) As Long
                                                   
Private Declare Function SetWindowPos Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE = -16
Private Const WS_CHILD As Long = &H400000
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const GWL_WNDPROC As Long = -4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const WS_EX_DLGMODALFRAME As Long = &H1&

Public Function HideTitleBar(ByVal window_caption As String)
    Dim window_info As Long
    Dim window_hwnd As Long
    
    window_hwnd = FindWindowA(vbNullString, window_caption)
    If window_hwnd <> 0 Then
        window_info = GetWindowLong(window_hwnd, GWL_STYLE)
        window_info = window_info And (Not WS_CHILD)
    
        Call SetWindowLong(window_hwnd, GWL_STYLE, window_info)
        Call DrawMenuBar(window_hwnd)
    End If
End Function

Public Function GoogleTalkWindow(ByVal window_caption As String)
    Dim window_info As Long
    Dim window_hwnd As Long
    
    window_hwnd = FindWindowA(vbNullString, window_caption)
    If window_hwnd <> 0 Then
        window_info = GetWindowLong(window_hwnd, GWL_STYLE)
        window_info = window_info And (Not WS_CHILD) And (Not WS_BORDER)
        
        ' changing the window style
        Call SetWindowLong(window_hwnd, GWL_STYLE, window_info)
        
        ' changing the extended style
        window_info = GetWindowLong(window_hwnd, GWL_EXSTYLE)
        window_info = window_info And (Not WS_EX_DLGMODALFRAME)
        Call SetWindowLong(window_hwnd, GWL_EXSTYLE, window_info)
        
        Call SetWindowPos(window_hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE + SWP_FRAMECHANGED)
        'Call DrawMenuBar(window_hwnd)
    End If
End Function
