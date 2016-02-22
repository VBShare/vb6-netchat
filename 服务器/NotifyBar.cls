VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NotifyBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'托盘模块
Option Explicit

Private Type POINTAPIs
    x As Long
    y As Long
End Type

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, Optional ByVal Length As Long = 4)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As NotifyBarActions, pnid As NOTIFYICONDATA) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPIs) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Enum NotifyClickClass
    NCL_MouseMove = 0
    NCL_DoubleClick = 1
    NCL_LeftButtonClick = 2
    NCL_RightButtonClick = 3
End Enum

Enum NotifyMessageIcons
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_GUID = &H4
End Enum

Private Enum NotifyBarActions
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIM_SETFOCUS = &H3
    NIM_SETVERSION = &H4
End Enum

Private Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Private Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const NIF_STATE = &H8
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10

Private Const WM_NOTIFYMESSAGE = 11122

Event NotifyClick(NClickClass As NotifyClickClass, ByVal x As Long, ByVal y As Long)

Private Type ThisClassSet
    
    s_wm_TaskBarCreated        As Long
    
    n_HandleOfWindow        As Long
    s_AddressOfWndProc      As Long
    
    n_NotifyData            As NOTIFYICONDATA
    
    n_ClickSign             As Boolean
    
End Type
Dim PG                      As ThisClassSet
Dim LinkProc()              As Long

Private Sub MsgHook(Result As Long, ByVal cHwnd As Long, ByVal Message As Long, ByVal wParam As Long, ByVal lParam As Long)

    '子类化接口过程
    
    Const WM_MOUSEMOVE = &H200
    Const WM_LBUTTONDBLCLK = &H203
    Const WM_LBUTTONDOWN = &H201
    Const WM_LBUTTONUP = &H202
    Const WM_RBUTTONDOWN = &H204
    Const WM_RBUTTONUP = &H205
    
    Dim Pss As POINTAPIs
    
    Result = CallWindowProc(PG.s_AddressOfWndProc, cHwnd, Message, wParam, lParam)
    
    If Message = PG.n_NotifyData.uCallbackMessage Then
    
        GetCursorPos Pss
        
        Select Case lParam
                Case WM_MOUSEMOVE: RaiseEvent NotifyClick(NCL_MouseMove, Pss.x, Pss.y)
                Case WM_LBUTTONDBLCLK: RaiseEvent NotifyClick(NCL_DoubleClick, Pss.x, Pss.y)
                Case WM_LBUTTONDOWN, WM_RBUTTONDOWN: PG.n_ClickSign = True
                Case WM_LBUTTONUP
                    If PG.n_ClickSign Then RaiseEvent NotifyClick(NCL_LeftButtonClick, Pss.x, Pss.y)
                    PG.n_ClickSign = False
                    
                Case WM_RBUTTONUP
                    SetForegroundWindow cHwnd
                    If PG.n_ClickSign Then RaiseEvent NotifyClick(NCL_RightButtonClick, Pss.x, Pss.y)
                    PG.n_ClickSign = False
                    
        End Select
        
    ElseIf Message = PG.s_wm_TaskBarCreated Then
        
        If PG.n_NotifyData.cbSize Then
            
            PG.n_NotifyData.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
            Shell_NotifyIcon NIM_ADD, PG.n_NotifyData
            
        End If
    End If

End Sub

Private Function GetWndProcAddress(ByVal SinceCount As Long) As Long
'   地址指针 ＝ GetWndProcAddress( SinceCount = 所有公用过程(例: Function,Sub,Property,公共变量) 总和 + 第 N 个 Private function/sub   )
'// 本例是取 MsgHook 函数地址, SinceCount 取值 12 ,可自行依例验证
    Dim mePtr As Long
    Dim jmpAddress As Long
    mePtr = ObjPtr(Me)
    CopyMemory jmpAddress, ByVal mePtr, 4
    CopyMemory jmpAddress, ByVal jmpAddress + (SinceCount - 1) * 4 + &H1C, 4

    ReDim LinkProc(10)
    LinkProc(0) = &H83EC8B55
    LinkProc(1) = &HFC8B14EC
    LinkProc(2) = &H56FC758D
    LinkProc(3) = &H3308758D
    LinkProc(4) = &HFC04B1C9
    LinkProc(5) = &HFF68A5F3
    LinkProc(6) = &HB8FFFFFF
    LinkProc(7) = &HFFFFFFFF
    LinkProc(8) = &H48BD0FF
    LinkProc(9) = &H10C2C924
    
    CopyMemory ByVal VarPtr(LinkProc(5)) + 3, mePtr, 4
    CopyMemory ByVal VarPtr(LinkProc(7)), jmpAddress, 4
    GetWndProcAddress = VarPtr(LinkProc(0))
    VirtualProtect ByVal VarPtr(LinkProc(0)), 44, &H40, mePtr
End Function

Private Sub Class_Initialize()
    
    PG.n_HandleOfWindow = CreateWindowEx(0, "button", vbNullString, &H40010F00, 0, 0, 10, 10, GetDesktopWindow&, ByVal 10&, App.hInstance, ByVal 0&)
    
    PG.n_NotifyData.uID = 1
    PG.n_NotifyData.hwnd = PG.n_HandleOfWindow
    PG.n_NotifyData.uCallbackMessage = WM_NOTIFYMESSAGE
    PG.s_AddressOfWndProc = SetWindowLong(ByVal PG.n_HandleOfWindow, ByVal -4&, ByVal GetWndProcAddress(12))
    PG.s_wm_TaskBarCreated = RegisterWindowMessage("TaskbarCreated")
End Sub

Private Sub Class_Terminate()

    Debug.Print "Start !NotifyBar Terminate ...", Timer

    SetWindowLong PG.n_HandleOfWindow, -4&, PG.s_AddressOfWndProc
    NotifyBoxVisible = False
    DestroyWindow PG.n_HandleOfWindow
    
    Debug.Print "NotifyBar Terminate ... end ", Timer
    
End Sub

Public Property Get Icon() As Long
    '// Notify Icon
    Icon = PG.n_NotifyData.hIcon
End Property

Public Property Let Icon(ByVal vNewValue As Long)
    
    With PG.n_NotifyData
        .uFlags = NIF_ICON Or .uFlags
        .hIcon = vNewValue
    End With
    
    If PG.n_NotifyData.cbSize Then
        PG.n_NotifyData.uFlags = NIF_ICON
        Shell_NotifyIcon NIM_MODIFY, PG.n_NotifyData
    End If
    
End Property

Public Property Get HandleOfNotifybar() As Long
    '// Notify hWnd
    HandleOfNotifybar = PG.n_NotifyData.hwnd
End Property

Public Property Let HandleOfNotifybar(ByVal vNewValue As Long)
    
    PG.n_NotifyData.hwnd = vNewValue
    If PG.n_NotifyData.cbSize Then _
            Shell_NotifyIcon NIM_MODIFY, PG.n_NotifyData
    
End Property

Public Property Get CallBackMessage() As Long
    '// Notify uCallbackMessage
    CallBackMessage = PG.n_NotifyData.uCallbackMessage
End Property

Public Property Let CallBackMessage(ByVal vNewValue As Long)
    
    PG.n_NotifyData.uFlags = NIF_MESSAGE Or PG.n_NotifyData.uFlags
    PG.n_NotifyData.uCallbackMessage = vNewValue
    
    If PG.n_NotifyData.cbSize Then
        PG.n_NotifyData.uFlags = NIF_MESSAGE
        Shell_NotifyIcon NIM_MODIFY, PG.n_NotifyData
    End If
    
End Property

Public Property Get ToolTipText() As String
    '// ToolTips
    ToolTipText = PG.n_NotifyData.szTip
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
    
    With PG.n_NotifyData
        .uFlags = NIF_TIP Or .uFlags
        .szTip = vNewValue & vbNullChar
    End With
    
    If PG.n_NotifyData.cbSize Then
        PG.n_NotifyData.uFlags = NIF_TIP
        Shell_NotifyIcon NIM_MODIFY, PG.n_NotifyData
    End If
    
End Property

Sub NotifyMsgBox(ByVal strMessage As String, Optional ByVal strTitle As String = "Application Notify", Optional ByVal NotifyMessageIcon As NotifyMessageIcons = NIIF_NONE)
    
    '// Notify MessageBox
    With PG.n_NotifyData
        .uFlags = NIF_INFO Or .uFlags
        .szInfo = strMessage & vbNullChar
        .szInfoTitle = strTitle & vbNullChar
        .dwInfoFlags = NotifyMessageIcon
    End With
    
    If PG.n_NotifyData.cbSize Then
        PG.n_NotifyData.uFlags = NIF_INFO
        Shell_NotifyIcon NIM_MODIFY, PG.n_NotifyData
    End If
End Sub

Public Property Get NotifyBoxVisible() As Boolean
    '// Notify visible
    If PG.n_NotifyData.cbSize Then NotifyBoxVisible = True
End Property

Public Property Let NotifyBoxVisible(ByVal vNewValue As Boolean)
    
    If vNewValue Then
        CallBackMessage = PG.n_NotifyData.uCallbackMessage
        PG.n_NotifyData.cbSize = Len(PG.n_NotifyData)
        Shell_NotifyIcon NIM_ADD, PG.n_NotifyData
        
    Else
        PG.n_NotifyData.uFlags = 0
        Shell_NotifyIcon NIM_DELETE, PG.n_NotifyData
        PG.n_NotifyData.cbSize = 0
    End If
    
End Property

