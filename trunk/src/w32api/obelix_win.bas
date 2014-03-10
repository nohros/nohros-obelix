Attribute VB_Name = "obelix_win"
Option Explicit

Private Const kDirectCOMDllName As String = "DirectCOM.dll"

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

Private Declare Function GetWindow Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByVal wCmd As Long) As Long
     
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" ( _
     ByVal hwnd As Long, _
     ByVal lpClassName As String, _
     ByVal nMaxCount As Long) As Long
     
Private Declare Function EnumChildWindows Lib "user32.dll" ( _
     ByVal hWndParent As Long, _
     ByVal lpEnumFunc As Long, _
     ByVal lparam As Long) As Long
     
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long
     
Private Declare Function SetParent Lib "user32.dll" ( _
     ByVal hWndChild As Long, _
     ByVal hWndNewParent As Long) As Long
     
Private Declare Function GetParent Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

Private Declare Function GetInstanceEx Lib "DirectCom" ( _
    StrPtr_FName As Long, _
    StrPtr_ClassName As Long, _
    Optional ByVal UseAlteredSearchPath As Boolean = True) As Object

Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" ( _
     ByVal lpLibFileName As String) As Long
     
Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" ( _
     ByVal lpModuleName As String) As Long
     
Declare Function GetAncestor Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByVal gaFlags As Long) As Long

Private Declare Function QueryPerformanceFrequency& Lib "kernel32.dll" (x@)
Private Declare Function QueryPerformanceCounter& Lib "kernel32.dll" (x@)

Private last_child_hwnd As Long

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const GW_HWNDNEXT As Long = 2
Private Const GW_CHILD As Long = 5
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE = -16
Private Const GWL_WNDPROC As Long = -4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const WS_MAXIMIZEBOX As Long = &H10000

' Window style and extented style
Private Const WS_POPUP As Long = &H80000000
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_CHILD As Long = &H400000
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_CAPTION As Long = &HC00000
Const WS_EX_TRANSPARENT As Long = &H20&

Private Const LWA_ALPHA As Long = &H2

Public Enum WindowStyle
    wsMaximizeBox = WS_MAXIMIZEBOX
    wsLayered = WS_EX_LAYERED
    wsPopUp = WS_POPUP
    wsChild = WS_CHILD
End Enum

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
    Dim window_hwnd As Long
    
    window_hwnd = FindWindowA(vbNullString, window_caption)
    If window_hwnd <> 0 Then
        GoogleTalkWindowHwnd window_hwnd
    End If
End Function

Public Function GoogleTalkWindowHwnd(ByVal window_hwnd As Long)
    Dim window_info As Long
    
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
End Function

'**
'* Removes the specified style from the specified window
'*
'* @param The HWND of the window to remove the style.
'* @param The style to be removed from the window class.
Public Function RemoveStyle(ByVal window_hwnd, ByVal style_to_remove As WindowStyle)
    Dim window_style As Long
    
    ' get the window extended style information
    window_style = GetWindowLong(window_hwnd, GWL_STYLE)
    
    ' remove the WS_BORDER style from the specified
    If window_style <> 0 Then
        window_style = window_style And (Not style_to_remove)
    End If
    
    ' commit the style changes
    Call SetWindowLong(window_hwnd, GWL_STYLE, window_style)
End Function

'**
'* Removes the specified extended style from the specified window
'*
'* @param The HWND of the window to remove the style.
'* @param The extended style to be removed from the window class.
Public Function RemoveExtendedStyle(ByVal window_hwnd, ByVal style_to_remove As WindowStyle)
    Dim window_style As Long
    
    ' get the window style information
    window_style = GetWindowLong(window_hwnd, GWL_EXSTYLE)
    
    ' remove the WS_BORDER style from the specified
    If window_style <> 0 Then
        window_style = window_style And (Not style_to_remove)
    End If
    
    ' commit the style changes
    Call SetWindowLong(window_hwnd, GWL_EXSTYLE, window_style)
End Function

'**
'* Add the specified extended style to the given window
'*
'* @param The handle of the window to set the style
'* @param The style to add to the window
Public Function SetExtendedStyle(ByVal window_hwnd As Long, ByVal style_to_set As WindowStyle)
    Dim window_style As Long
    
    'get the window style information
    window_style = GetWindowLong(window_hwnd, GWL_EXSTYLE)
    
    ' add the specified style to the window
    If window_style <> 0 Then
        window_style = window_style Or style_to_set
    End If
    
    'commit the style changes
    Call SetWindowLong(window_hwnd, GWL_EXSTYLE, window_style)
End Function

'**
'* Add the specified window style to the given window
'*
'* @param The handle of the window to set the style
'* @param The style to add to the window
Public Function SetStyle(ByVal window_hwnd As Long, ByVal style_to_set As WindowStyle)
    Dim window_style As Long
    
    'get the window style information
    window_style = GetWindowLong(window_hwnd, GWL_STYLE)
    
    ' add the specified style to the window
    If window_style <> 0 Then
        window_style = window_style Or style_to_set
    End If
    
    'commit the style changes
    Call SetWindowLong(window_hwnd, GWL_STYLE, window_style)
End Function

Public Function GetBrowserWindowCallback(ByVal child_hwnd As Long, ByVal lparam As Long) As Long
    Dim class_name_buffer As String
    Dim class_name As String
    Dim ret As Long
    Dim continue_enumerating As Long
    
    continue_enumerating = 1
    
    ' create the buffer to hold the class name
    class_name_buffer = Space$(256)
    
    ret = GetClassName(child_hwnd, class_name_buffer, 256)
    If ret > 0 Then
        class_name = StrConv(Mid(class_name_buffer, 1, ret), vbLowerCase)
        If class_name = "shell embedding" Then
            continue_enumerating = 0
            last_child_hwnd = child_hwnd
        End If
    End If
    
    GetBrowserWindowCallback = continue_enumerating
End Function

'**
'* Gets the HWND for first webbrowser that is child of the specified windows
'*
'* @param The HWND of the parent window.
'*
'* @return The HWND of a window that is child of the specified parent window and whose
'*         class name is [shell embedding]
Public Function GetBrowserWindow(hwnd_parent As Long) As Long
    Dim RetVal As Long
    Dim result As Long
    Dim hwnd_child As Long
    Dim class_name As String
    
    last_child_hwnd = 0
    
    EnumChildWindows hwnd_parent, AddressOf GetBrowserWindowCallback, ByVal 0&
    
    GetBrowserWindow = last_child_hwnd
End Function

Public Sub ToggleTransparency(ByVal window_hwnd As Long, ByVal alpha As Long)
    Dim parent_hwnd As Long
    Dim window_exstyle As Long
    
    parent_hwnd = GetAncestor(window_hwnd, 1)
    
    SetParent window_hwnd, 0&
    SetStyle window_hwnd, WS_CAPTION
    
    window_exstyle = GetWindowLong(window_hwnd, GWL_EXSTYLE)
    If (window_exstyle And WS_EX_LAYERED) = 0 Then
        SetExtendedStyle window_hwnd, wsLayered
        SetExtendedStyle window_hwnd, WS_EX_TRANSPARENT
        SetLayeredWindowAttributes window_hwnd, 0, alpha, LWA_ALPHA
    Else
        RemoveExtendedStyle window_hwnd, wsLayered
        RemoveExtendedStyle window_hwnd, WS_EX_TRANSPARENT
    End If
        
    RemoveStyle window_hwnd, WS_CAPTION
    SetParent window_hwnd, parent_hwnd
End Sub

'**
'* Creates and returns a reference to a COM object using the specified manifest file.
'* <p>
'* This method is used to create objects without using the windows registry(registration free method).
'*
'* @param com_class The name of the COM class of the object to create.
'* @param com_manifest The manifest file that describes the object to create.
Public Function CreateRegFreeObject(ByVal com_class_name As String, ByVal com_dll_path As String, ByRef com_object As Object) As Boolean
    Dim direct_com_handle As Long
    Dim direct_com_dll_path As String
    Dim result As Boolean
    
    result = False
    
    On Error GoTo Finally
    
    direct_com_dll_path = obelix_io.GetBinPathFor(kDirectCOMDllName)
    
    ' check if the DirectCom.dll module is already loaded
    direct_com_handle = GetModuleHandle(direct_com_dll_path)
    If direct_com_handle = 0 Then
        ' the DirectCOM dll is not loaded, load it.
        direct_com_handle = LoadLibrary(direct_com_dll_path)
        If direct_com_handle = 0 Then
            Err.Raise vbError + 500, "CreateRegFreeObject", "The DirectCOM dll could not be loaded."
            GoTo Finally
        End If
    End If
    
    Set com_object = GetInstanceEx(StrPtr(com_dll_path), StrPtr(com_class_name), True)
    
    result = True
    
    GoTo Finally
    
Catch:
    Set com_object = Nothing
    result = False
    
    LogError Err
    
    ' propagate the erro to the caller
    'Err.Raise Err.number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
Finally:
    CreateRegFreeObject = result
End Function

Public Function HPTimer#()
    Dim x@: Static Frq@
    
    If Frq = 0 Then QueryPerformanceFrequency Frq
    If QueryPerformanceCounter(x) Then HPTimer = x / Frq
End Function

Public Function Timing(Optional ByVal Start As Boolean) As Double
    Static T#
    
    If Start Then T = HPTimer: Exit Function
    Timing = HPTimer - T
End Function

Public Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    FindWindow = FindWindowA(lpClassName, lpWindowName)
End Function
