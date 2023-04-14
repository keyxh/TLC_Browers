Attribute VB_Name = "head_file"

Public Declare Function URLDownloadToFile Lib "urlmon.dll" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long



Public Declare Function ReleaseCapture Lib "user32" () As Long '无窗体解锁
'''无边框窗体移动'''
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'''等待程序运行结束'''
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'''无边框窗体缩放'''

Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As rect) As Long
'设置焦点
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long







'''常量声明'''

Public Const WS_EX_MDICHILD As Long = &H40&
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_DISABLED = &H8000000
Public Const WS_GROUP = &H20000
Public Const WS_DIGFRAME = &H400000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = &H40000
Public Const WM_SETFOCUS = &H7



Public Const SWP_NOMOVE = &H2 '不移动窗体
Public Const SWP_NOSIZE = &H1 '不改变窗体尺寸
Public Const Flag = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1 '窗体总在最前面
Public Const HWND_NOTOPMOST = -2 '窗体不在最前面
Public Const theScreen = 0
Public Const theForm = 1
Public Const SWP_NOZORDER = &H4
Public Const SWP_DRAWFRAME = &H20




'''自定义类型'''

Public Type rect
    left As Long
    top As Long
    Right As Long
    Botton As Long
End Type















