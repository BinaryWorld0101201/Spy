Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWindow As Long, ByVal clval As Long, ByVal alpha As Byte, ByVal flago As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hnd As Long, ByVal OldIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWindow As Long, Margin As MARGINS) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWindow As Long, ByVal OldIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWindow As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWindow As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWindow As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWindow As Long, lpRect As RECT) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWindow As Long, lpdwProcessId As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWindow As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWindow As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWindow As Long, ByVal fEnable As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWindow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function Module32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWindow As Long, ByVal hDC As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWindow As Long, ByVal wCmd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWindow As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Public Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWindow As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetPriorityClass Lib "kernel32" (ByVal hProcess As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" (lpShellInfo As SHELLEXECUTEINFO) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

' For Aero Glass
Public Declare Function SetWindowCompositionAttribute Lib "user32" (ByVal hWnd As Long, ByRef AttrData As WCAD) As Long
Public Type WCAD
    Attr As Long
    pData As Long
    cbData As Long
End Type
Public Type ACCENT_POLICY
    AccentState As Long
    AccentFlags As Long
    GradientColor As Long
    AnimationId As Long
End Type
Public Const ACCENT_ENABLE_BLURBEHIND = 3
Public Const WCA_ACCENT_POLICY = 19

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWindow As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 1024
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type MARGINS
    m_Left As Long
    m_Right As Long
    m_Top As Long
    m_Buttom As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const WM_SETTEXT = &HC
Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const PS_SOLID = 0
Public Const TH32CS_SNAPMODULE = &H8
Public Const SC_MONITORPOWER = &HF170&
Public Const WM_SYSCOMMAND = &H112
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_RESTORE = 9
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWNORMAL = 1
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const LWA_COLORKEY = &H1&
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const WM_CLOSE = &H10
Public Const DT_Mid = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
Public Const SWP_No = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Public Const LB_FINDSTRING = &H18F


Public Const WS_VSCROLL = &H200000
Public Const WS_VISIBLE = &H10000000
Public Const WS_THICKFRAME = &H40000
Public Const WS_TABSTOP = &H10000
Public Const WS_SYSMENU = &H80000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_POPUP = &H80000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_HSCROLL = &H100000
Public Const WS_GROUP = &H20000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_LAYERED = &H80000
Public Const WS_DLGFRAME = &H400000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CHILD = &H40000000
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_CENTER = &H1&
Public Const ES_LEFT = &H0&
Public Const ES_LOWERCASE = &H10&
Public Const ES_MULTILINE = &H4&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_PASSWORD = &H20&
Public Const ES_READONLY = &H800&
Public Const ES_RIGHT = &H2&
Public Const ES_UPPERCASE = &H8&
Public Const ES_WANTRETURN = &H1000&

Public Const HIGH_PRIORITY_CLASS = &H80
Public Const IDLE_PRIORITY_CLASS = &H40
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const REALTIME_PRIORITY_CLASS = &H100
Public Const BELOW_NORMAL = 16384
Public Const ABOVE_NORMAL = 32768
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Public Const MAX_PATH = 260


