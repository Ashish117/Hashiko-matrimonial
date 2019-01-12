Attribute VB_Name = "Module1"
Public Const WS_CHILD As Long = &H40000000
Public Const WS_VISIBLE As Long = &H10000000


Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER



Public Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Public Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Public Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Public Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Public Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25







Public Declare Function capCreateCaptureWindow _
    Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
         (ByVal lpszWindowName As String, ByVal dwStyle As Long _
        , ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long _
        , ByVal nHeight As Long, ByVal hwndParent As Long _
        , ByVal nID As Long) As Long







Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long _
        , ByVal wParam As Long, ByRef lParam As Any) As Long


