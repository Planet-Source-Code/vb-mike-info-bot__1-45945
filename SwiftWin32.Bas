Attribute VB_Name = "Module1"
Public color1 As String, color2 As String, DialerLogin As String
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Public Const MOUSEEVENTF_MIDDLEUP = &H40
    Public Const MOUSEEVENTF_RIGHTDOWN = &H8
    Public Const MOUSEEVENTF_RIGHTUP = &H10
    Public Const MOUSEEVENTF_MOVE = &H1
Public Enum enmFilePath
    FileName
    Path
    Drive
End Enum


'MikesTooLz win 32 bit bas 99
'The ultimate advanced vb4, 5, 6 programming bas
'includes winsock dll declarations anbd subs, windows
'command decs and many AOL 4.0 subs
'75% of AOL 4.0 subs and Decs thanks to izekial


'|---Menu Picture Declarations---|

Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
    Public Const MF_BITMAP = &H4&
Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    FState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    End Type
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
    Public Const MIIM_ID = &H2
    Public Const MIIM_TYPE = &H10
    Public Const MFT_STRING = &H0&
  '-----------
Declare Function GetWinFlags Lib "Kernel" () As Long
'***************************************************************

'Windows API/Global Declarations for :cRegistry

'***************************************************************




Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
    End Type


Type FIIME
    dwLowDateTime As Long
    dwHighDateTime As Long
    End Type
    '----Registry Function Prototypes



Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FIIME) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegDeeKey Lib "advapi32.dll" Alias "RegDeeKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeeValue Lib "advapi32.dll" Alias "RegDeeValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'|---- Winsock SMTP Mail Attachment Dec's----|


Private Const LOCALE_SSHORTDATE = &H1F
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&

Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Const FD_SETSIZE = 64
Type fd_set
    fd_count As Integer
    fd_array(FD_SETSIZE) As Integer
End Type

Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const hostent_size = 16

Type servent
    s_name As Long
    s_aliases As Long
    s_port As Integer
    s_proto As Long
End Type
Public Const servent_size = 14

Type protoent
    p_name As Long
    p_aliases As Long
    p_proto As Integer
End Type
Public Const protoent_size = 10

Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17

Public Const INADDR_NONE = &HFFFF
Public Const INADDR_ANY = &H0

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const sockaddr_size = 16
Public saZero As sockaddr


Public Const WSA_DESCRIPTIONLEN = 256
Public Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Public Const WSA_SYS_STATUS_LEN = 128
Public Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1

Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2

Public Const MAXGETHOSTSTRUCT = 1024

Public Const AF_INET = 2
Public Const PF_INET = 2

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type
' Windows Sockets definitions of regular Microsoft C error constants
Global Const WSAEINTR = 10004
Global Const WSAEBADF = 10009
Global Const WSAEACCES = 10013
Global Const WSAEFAULT = 10014
Global Const WSAEINVAL = 10022
Global Const WSAEMFILE = 10024
' Windows Sockets definitions of regular Berkeley error constants
Global Const WSAEWOULDBLOCK = 10035
Global Const WSAEINPROGRESS = 10036
Global Const WSAEALREADY = 10037
Global Const WSAENOTSOCK = 10038
Global Const WSAEDESTADDRREQ = 10039
Global Const WSAEMSGSIZE = 10040
Global Const WSAEPROTOTYPE = 10041
Global Const WSAENOPROTOOPT = 10042
Global Const WSAEPROTONOSUPPORT = 10043
Global Const WSAESOCKTNOSUPPORT = 10044
Global Const WSAEOPNOTSUPP = 10045
Global Const WSAEPFNOSUPPORT = 10046
Global Const WSAEAFNOSUPPORT = 10047
Global Const WSAEADDRINUSE = 10048
Global Const WSAEADDRNOTAVAIL = 10049
Global Const WSAENETDOWN = 10050
Global Const WSAENETUNREACH = 10051
Global Const WSAENETRESET = 10052
Global Const WSAECONNABORTED = 10053
Global Const WSAECONNRESET = 10054
Global Const WSAENOBUFS = 10055
Global Const WSAEISCONN = 10056
Global Const WSAENOTCONN = 10057
Global Const WSAESHUTDOWN = 10058
Global Const WSAETOOMANYREFS = 10059
Global Const WSAETIMEDOUT = 10060
Global Const WSAECONNREFUSED = 10061
Global Const WSAELOOP = 10062
Global Const WSAENAMETOOLONG = 10063
Global Const WSAEHOSTDOWN = 10064
Global Const WSAEHOSTUNREACH = 10065
Global Const WSAENOTEMPTY = 10066
Global Const WSAEPROCLIM = 10067
Global Const WSAEUSERS = 10068
Global Const WSAEDQUOT = 10069
Global Const WSAESTALE = 10070
Global Const WSAEREMOTE = 10071
' Extended Windows Sockets error constant definitions
Global Const WSASYSNOTREADY = 10091
Global Const WSAVERNOTSUPPORTED = 10092
Global Const WSANOTINITIALISED = 10093
Global Const WSAHOST_NOT_FOUND = 11001
Global Const WSATRY_AGAIN = 11002
Global Const WSANO_RECOVERY = 11003
Global Const WSANO_DATA = 11004
Global Const WSANO_ADDRESS = 11004
'---ioctl Constants
    Public Const FIONREAD = &H8004667F
    Public Const FIONBIO = &H8004667E
    Public Const FIOASYNC = &H8004667D


#If Win32 Then
'---Windows System Functions

    Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Src As Any, ByVal cb&)
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'---async notification constants
    Public Const SOL_SOCKET = &HFFFF&
    Public Const SO_LINGER = &H80&
    Public Const FD_READ = &H1&
    Public Const FD_WRITE = &H2&
    Public Const FD_OOB = &H4&
    Public Const FD_ACCEPT = &H8&
    Public Const FD_CONNECT = &H10&
    Public Const FD_CLOSE = &H20&
'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
    Public Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal s As Long, sName As sockaddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, FROM As sockaddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function send Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, to_addr As sockaddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
    Public Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Sub WSASetLastError Lib "wsock32.dll" (ByVal iError As Long)
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal number As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
#End If


'SOME STUFF I ADDED
Public MySocket%
Public SockReadBuffer$
Public Const WSA_NoName = "Unknown"
Public WSAStartedUp As Boolean     'Flag to keep track of whether winsock WSAStartup wascalled

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeeObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeeMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal DWreserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long




'Window Messages
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &HF012
Public Const WM_gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CLEAR = &H303
Public Const WM_DRAWITEM = &H2B
Public Const WM_PAINT = &HF
Public Const WM_ERASEBKGND = &H14
Public Const WM_NCPAINT = &H85



'Combo Box Functions
Public Const CB_DEESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETCOUNT = &H146

'hWnd Functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'Show Window Functions
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1

'Sound Functions
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SND_LOOP = &H8

'Screen Saver Function
Public Const SPI_SCREENSAVERRUNNING = 97

'Get Window Word Functions
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

'Virtual Key Statements
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
'Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

'Phader Color Presets
Public Const COLOR_RED = &HFF&
Public Const COLOR_GREEN = &HFF00&
Public Const COLOR_BLUE = &HFF0000
Public Const COLOR_YELLOW = &HFFFF&
Public Const COLOR_WHITE = &HFFFFFE
Public Const COLOR_BLACK = &H0&
Public Const COLOR_PEACH = &HC0C0FF
Public Const COLOR_PURPLE = &HFF00FF
Public Const COLOR_GREY = &HC0C0C0
Public Const COLOR_PINK = &HFF80FF
Public Const COLOR_TURQUOISE = &HC0C000
Public Const COLOR_LIGHTBLUE = &HFF8080
Public Const COLOR_ORANGE = &H80FF&

'Processor Types
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

'Menu Functions
Public Const MF_DEE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_POPUP = &H10&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&

'Key Presets
Public Const ENTER_KEY = 13

'Button Messages
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

'List Box Functions
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DEESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETtext = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

'Notify Icon Functions
Public Const NIM_ADD = &H0
Public Const NIM_DEE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

'Edit Window Messages
Public Const EM_REPLACESEL = &HC2
Public Const EM_SETSEL = &HB1

'Dev Mode Const's
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000

'Windows Version Functions
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Byte Functions
Public Const MAX_DEFAULTCHAR = 2
Public Const MAX_LEADBYTES = 12

'winsck functions
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128

Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1

Public Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Dim DevM As DEVMODE

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        DWreserved As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Type CPINFO
        MaxCharSize As Long
        DefaultChar(MAX_DEFAULTCHAR - 1) As Byte
        LeadByte(MAX_LEADBYTES - 1) As Byte
End Type

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type COLORRGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Public Type POINTAPI
    x As Long
    Y As Long
End Type


Public chatline As String

Declare Function MoveWindow Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal x As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long

    Public Const LB_INITSTORAGE = &H1A8

    Public Const WM_SETREDRAW = &HB
    Public Const WM_VSCROLL = &H115
    Public Const SB_BOTTOM = 7
Declare Function GetLogicalDrives Lib "kernel32" () As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    Public Const INVALID_HANDLE_VALUE = -1
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

    Public Const MaxLFNPath = 260
Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FIIME
    ftLastAccessTime As FIIME
    ftLastWriteTime As FIIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MaxLFNPath
    cShortFileName As String * 14
    End Type

Type Com
    Reply As String
    BackCommand As String
End Type
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Sub clicksend()
Dim AOL As Long, MDI As Long, ErRor As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String
DoEvents
        OpenForward& = FindForwardWindow
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&

        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        timeout (1)
    
End Sub

Function EjectCD()
mciSendString "set cd door open", 0, 0, 0
End Function
Public Function deelinefromstring(strstring As String, lngline As Long) As String
    Dim lnglinecount As Long, index As Long, strText As String
     lnglinecount& = getstringlinecount(strstring)
    For index& = 1& To lnglinecount&
        If index& <> lngline& Then
            strText$ = strText$ & getlinefromstring(strstring, index&) & vbCrLf$
        End If
    Next index&
    deelinefromstring$ = strText$
End Function

Public Sub deenewmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, lngbutton As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
             lngbutton& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub

Sub CPanel_Open()
'Sub declared by Tech
Call ControlPanels("C:\Windows\Control.Exe")
End Sub

Sub cursor_move()
Call SetCursorPos(100, 50)
End Sub

Public Sub Text_Encrypt(TheText)
'Encrypt Data
TheText.Text = CryptString2(TheText.Text, True)
'this is done in a Public Function.
'CryptString2 = Function Name
'text5.text = Data to be Encrypted
'True = Not Encrypted Yet
'False = Encrypted, must be decrypted
End Sub


Public Sub Text_Decrypt(TheText)
'decrypt Data
TheText.Text = CryptString2(TheText.Text, False)
'this is done in a Public Function.
'CryptString2 = Function Name
'True = Not Encrypted Yet
'False = Encrypted, must be decrypted
End Sub
Function decrypt(Txt As String, PW As String) As String
Dim pwcounter
Dim x
Dim Curter
Dim Text$
Dim pwter
Dim Outt
Dim Outter
Dim ascPW
Dim Combined
Dim Ascter
If Len(Txt$) = 0 Then Exit Function
If Len(PW$) = 0 Then Exit Function
pwcounter = 1
For x = 1 To Len(Txt$)
Curter = Mid(Txt$, x, 1)
Ascter = Asc(Curter)
pwter = Mid(PW$, pwcounter, 1)
pwcounter = pwcounter + 1
If pwcounter > Len(PW$) Then pwcounter = 1
ascPW = Asc(pwter)
Combined = Ascter - ascPW
If Combined < 1 Then
Outter = Combined + 255
Outt = Chr(Outter)
Else
Outt = Chr(Combined)
End If
Text$ = Text$ + Outt
DoEvents
Next x
decrypt = Text$
End Function
Sub emergency_bkup(Text2 As TextBox)
If Not file_IfDirectoryexists(Text2) Then
MsgBox "This directory does not exist.  PC Medic will now create the defult directory [c:\emergency]", 16, "ok"
MkDir "c:\emergency"
Text2 = "c:\emergency"
GoTo makeit
End If
makeit:
If Not file_ifileexists("c:\windows\win.ini") Then GoTo wincom
FileCopy ("c:\windows\win.ini"), (Text2 & "\win.ini")
wincom:
If Not file_ifileexists("c:\windows\win.com") Then GoTo comcom
FileCopy ("c:\windows\win.com"), (Text2 & "\win.com")
comcom:
If Not file_ifileexists("c:\command.com") Then GoTo execbat
FileCopy ("c:\command.com"), (Text2 & "\command.com")
execbat:
If Not file_ifileexists("c:\autoexec.bat") Then GoTo sysini
FileCopy ("c:\autoexec.bat"), (Text2 & "\autoexec.bat")
sysini:
If Not file_ifileexists("c:\windows\system.ini") Then GoTo mad
FileCopy ("c:\windows\system.ini"), (Text2 & "\system.ini")
mad:
If Not file_ifileexists("c:\autoexec.mad") Then GoTo config
FileCopy ("c:\autoexec.mad"), (Text2 & "\autoexec.mad")
config:
If Not file_ifileexists("c:\config.sys") Then GoTo boottxt
FileCopy ("c:\config.sys"), (Text2 & "\config.sys")
boottxt:
If Not file_ifileexists("c:\bootlog.txt") Then GoTo boot
FileCopy ("c:\bootlog.txt"), (Text2 & "\bootlog.txt")
boot:
If Not file_ifileexists("c:\bootlog.prv") Then GoTo msdossys
FileCopy ("c:\bootlog.txt"), (Text2 & "\bootlog.prv")
msdossys:
If Not file_ifileexists("c:\msdos.sys") Then GoTo sta
FileCopy ("c:\msdos.sys"), (Text2 & "\msdos.sys")
sta:
If Not file_ifileexists("c:\windows\command\start.exe") Then GoTo sys2
FileCopy ("c:\windows\command\start.exe"), (Text2 & "\start.exe")
sys2:
If Not file_ifileexists("c:\windows\command\sys.exe") Then GoTo done:
FileCopy ("c:\windows\command\start.exe"), (Text2 & "\sys.exe")
done:

End Sub


Public Sub forwardnewmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call opennewmailbox
    End If
    Call opennewmailitem(index&)
    Do: DoEvents
        Pause 1&
         lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getnewmailsubject(index&))
         lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
         lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
         lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
         lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
         lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
         lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
         lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
         lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
         lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
         lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
         lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
         lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
         lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
         lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
         lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
         lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
         lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
         lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
         lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
         lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
         lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
         lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
         lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
         lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
         lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
         lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
         lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        Pause 0.2
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deefwdtext = True Then
         strnewsubject$ = GetText(lngsubjectwin&)
         strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
         lngaol& = FindWindow("AOL Frame25", vbNullString)
         lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
         lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call Pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
     lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getnewmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
             lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
             lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub

Public Sub forwardoldmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call openoldmailbox
    End If
    Call openoldmailitem(index&)
    Do: DoEvents
        Pause 1&
         lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getoldmailsubject(index&))
         lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
         lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
         lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
         lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
         lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
         lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
         lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
         lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
         lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
         lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
         lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
         lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
         lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
         lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
         lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
         lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
         lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
         lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
         lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
         lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
         lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
         lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
         lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
         lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
         lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
         lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
         lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        Pause 0.2
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deefwdtext = True Then
         strnewsubject$ = GetText(lngsubjectwin&)
         strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
         lngaol& = FindWindow("AOL Frame25", vbNullString)
         lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
         lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call Pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
     lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getoldmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
             lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
             lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub

Public Sub formtotray(frmForm As Form)
    Dim systray As NOTIFYICONDATA
    With systray
         .cbSize = Len(systray)
         .uID = vbNull
         .hwnd = frmForm.hwnd
         .uCallbackMessage = WM_MOUSEMOVE
         .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         .hIcon = frmForm.Icon
         .szTip = frmForm.Caption
    End With
    Call Shell_NotifyIcon(NIM_ADD, systray)
    frmForm.Hide
End Sub

Public Sub formunloadright(frmForm As Form)
    Do: DoEvents
         frmForm.Left = frmForm.Left + 250&
    Loop Until frmForm.Left > Screen.Width
    Unload frmForm
End Sub

Public Sub forwardallflashmail(strwho As String, strmessage As String, deefwdtext As Boolean)
    Dim index As Long
    Call openflashmailbox
    For index& = 0& To getflashmailcount& - 1&
        Call forwardflashmailitem(True, strwho$, strmessage$, index&, deefwdtext): DoEvents
    Next index&
End Sub

Public Sub forwardallflashmailwithdogbar(strwho As String, strmessage As String, deefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call openflashmailbox
     progbar.Max = getflashmailcount& - 1&
    For index& = 0& To getflashmailcount& - 1&
         progbar.Value = index& - 1&
        Call forwardflashmailitem(True, strwho$, strmessage$, index&, deefwdtext): DoEvents
    Next index&
End Sub

Public Sub forwardallnewmail(strwho As String, strmessage As String, deefwdtext As Boolean)
    Dim index As Long
    Call opennewmailbox
    For index& = 0& To getnewmailcount& - 1&
        Call forwardnewmailitem(True, strwho$, strmessage$, index&, deefwdtext): DoEvents
    Next index&
End Sub

Public Sub forwardallnewmailwithdogbar(strwho As String, strmessage As String, deefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call opennewmailbox
     progbar.Max = getnewmailcount& - 1&
    For index& = 0& To getnewmailcount& - 1&
         progbar.Value = index& - 1&
        Call forwardnewmailitem(True, strwho$, strmessage$, index&, deefwdtext): DoEvents
    Next index&
End Sub

Public Sub forwardalloldmail(strwho As String, strmessage As String, deefwdtext As Boolean)
    Dim index As Long
    Call openoldmailbox
    For index& = 0& To getoldmailcount& - 1&
        Call forwardoldmailitem(True, strwho$, strmessage$, index&, deefwdtext): DoEvents
    Next index&
End Sub

Public Sub forwardalloldmailwithdogbar(strwho As String, strmessage As String, deefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call openoldmailbox
     progbar.Max = getoldmailcount& - 1&
    For index& = 0& To getoldmailcount& - 1&
         progbar.Value = index& - 1&
        Call forwardoldmailitem(True, strwho$, strmessage$, index&, deefwdtext): DoEvents
    Next index&
End Sub



Public Sub forwardflashmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call openflashmailbox
    End If
    Call openflashmailitem(index&)
    Do: DoEvents
        Pause 1&
         lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getflashmailsubject(index&))
         lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
         lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
         lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
         lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
         lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
         lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
         lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
         lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
         lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
         lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
         lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
         lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
         lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
         lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
         lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
         lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
         lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
         lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
         lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
         lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
         lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
         lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
         lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
         lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
         lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
         lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
         lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        Pause 0.3
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deefwdtext = True Then
         strnewsubject$ = GetText(lngsubjectwin&)
         strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
         lngaol& = FindWindow("AOL Frame25", vbNullString)
         lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
         lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call Pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
     lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getflashmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
             lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
             lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub

Public Sub FormDrag(frmForm As Form)
    Call ReleaseCapture
    Call SendMessage(frmForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Sub formfromtray(frmForm As Form)
    Dim systray As NOTIFYICONDATA
    With systray
         .cbSize = Len(systray)
         .hwnd = frmForm.hwnd
         .uID = vbNull
    End With
    Call Shell_NotifyIcon(NIM_DEE, systray)
    frmForm.Show
End Sub

Public Sub forminnerbevel(frmForm As Form, lngbevelwidth As Long)
     frmForm.ScaleMode = 3&
    Call formbevellines(frmForm, 0&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmForm, 1&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmForm, 2&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmForm, 3&, lngbevelwidth&, QBColor(15&))
End Sub


Public Sub formleaveright(frmForm As Form)
    Do: DoEvents
        frmForm.Left = frmForm.Left + 250&
    Loop Until frmForm.Left > Screen.Width
End Sub

Public Sub FormOnTop(frmForm As Form, OnTop As Boolean)
    If OnTop = True Then Call SetWindowPos(frmForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
    If OnTop = False Then Call SetWindowPos(frmForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Function get3colorbarsvalue(RedBar As Control, GreenBar As Control, BlueBar As Control) As Long
    get3colorbarsvalue& = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
End Function

Public Function getaolversion() As Long
    Dim lngaol As Long, lngtoolwin As Long, lngtoolico As Long
    Dim lngcombo As Long, lngcomboedit As Long
     lngaol& = FindWindow("aol frame25", "america  online")
     lngtoolwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
     lngtoolico& = FindWindowEx(lngtoolwin&, 0&, "_aol_toolbar", vbNullString)
     lngcombo& = FindWindowEx(lngtoolico&, 0&, "_aol_combobox", vbNullString)
     lngcomboedit& = FindWindowEx(lngcombo&, 0&, "edit", vbNullString)
    If lngcomboedit& = 0& Then
         getaolversion& = 3&
    ElseIf lngcomboedit& > 0& Then
         getaolversion& = 4&
    End If
End Function

Public Function GetCaption(lngwindow As Long) As String
    Dim strBuffer As String, lngtextlen As Long
     lngtextlen& = GetWindowTextLength(lngwindow&)
     strBuffer$ = String$(lngtextlen&, 0&)
    Call GetWindowText(lngwindow&, strBuffer$, lngtextlen& + 1&)
     GetCaption$ = strBuffer$
End Function

Public Function getchatline() As String
  Dim ChatWin As Long, strText As String, lngenter As Long, lngenter2 As Long
     ChatWin& = FindRoom&
     ChatWin& = FindWindowEx(ChatWin&, 0&, "RICHCNTL", vbNullString)
     strText$ = GetText(ChatWin&)
     lngenter& = InStr(strText$, Chr$(13&))
    Do: DoEvents
        If lngenter& <> 0& Then lngenter2& = lngenter&
         lngenter& = InStr(lngenter& + 1&, strText$, Chr$(13&))
    Loop Until lngenter& = 0&
     getchatline$ = Right$(strText$, Len(strText$) - lngenter2&)
End Function
Public Function getchatlinemsg() As String
    Dim strline As String, lngtab As Long
     strline$ = getchatline$
     lngtab& = InStr(strline$, Chr$(9&))
     getchatlinemsg$ = Right$(strline$, Len(strline$) - lngtab&)
End Function

Public Function getchatlinesn() As String
    Dim strline As String, lngtab As Long
     strline$ = getchatline$
     lngtab& = InStr(strline$, Chr$(9&))
     getchatlinesn$ = Left$(strline$, lngtab& - 3&)
End Function

Public Function getchatname() As String
     getchatname$ = GetCaption(FindRoom&)
End Function


Public Function GetChatText() As String
    Dim ChatWin As Long
     ChatWin& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
     GetChatText$ = GetText(ChatWin&)
End Function

Public Function getcomputerprocessor() As String
    'from frenzy3.bas by izekial(me)
    Dim sstr As String, myver As OSVERSIONINFO, mysys As SYSTEM_INFO
    Call GetSystemInfo(mysys)
    Select Case mysys.dwProcessorType
        Case PROCESSOR_INTEL_386
             sstr$ = "intel 386 dx"
        Case PROCESSOR_INTEL_486
             sstr$ = "intel 486 dx"
        Case PROCESSOR_INTEL_PENTIUM
             sstr$ = "intel pentium pro"
        Case PROCESSOR_MIPS_R4000
             sstr$ = "mips r-4000"
        Case PROCESSOR_ALPHA_21064
             sstr$ = "alpha 21064"
        Case Else
             sstr$ = "unknown processor"
    End Select
    If mysys.dwNumberOrfProcessors > 1 Then
         sstr$ = "multiple " & sstr$ & " processors"
    Else
         sstr$ = sstr$ & " processor"
    End If
     getcomputerprocessor$ = sstr$
End Function

Public Function getdiskinfo(thedrive As String) As String
    'from frenzy3.bas by izekial(me)
    Dim dl&, s$, spaceloc%, freebytes&, totalbytes&
    Dim sectorspercluster&, bytespersector&, numberoffreeclustors&
    Dim bytesfree&, bytesTotal&, percentfree&, totalnumberofclustors&
    Dim tmp1$, tmp2$, tmp3$, tmp4$, tmp5$, tmp6$, tmp7$
    If Right(thedrive$, 1) = "\" Then
         thedrive$ = Left(thedrive$, Len(thedrive$) - 1)
    End If
    If InStr(thedrive$, ":") Then
         thedrive$ = thedrive$
    Else
         thedrive$ = thedrive$ & ":"
    End If
     dl& = GetDiskFreeSpace(thedrive$, sectorspercluster, bytespersector, numberoffreeclustors, totalnumberofclustors)
     tmp1$ = "sectors per cluster :" & Format(sectorspercluster, "#,0")
     tmp2$ = "bytes per sector : " & Format(bytespersector, "#,0")
     tmp3$ = "number of free clusters : " & Format(numberoffreeclustors, "#,0")
     tmp4$ = "total number of clustors : " & Format(totalnumberofclustors, "#,0")
     tmp5$ = totalnumberofclustors * sectorspercluster * bytespersector
     tmp5$ = "total bytes : " & Format(tmp5$, "#,0")
     tmp6$ = numberoffreeclustors * sectorspercluster * bytespersector
     tmp6$ = "total free bytes : " & Format(tmp6$, "#,0")
     getdiskinfo$ = tmp1$ & vbCrLf & tmp2$ & vbCrLf & tmp3$ & vbCrLf & tmp4$ & vbCrLf & tmp5$ & vbCrLf & tmp6$
End Function

Public Sub formouterbevel(frmForm As Form, lngbevelwidth As Long)
     frmForm.ScaleMode = 3&
    Call formbevellines(frmForm, 0&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmForm, 1&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmForm, 2&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmForm, 3&, lngbevelwidth&, QBColor(8&))
End Sub
Public Sub formrattle(frmForm As Form)
    Dim index As Long, firstleft As Long
     firstleft& = frmForm.Left
    frmForm.Show
     frmForm.DrawMode = 2&
    For index& = 0& To 15&
         frmForm.Left = firstleft& + 25&
         frmForm.Left = frmForm.Left - 50&
    Next index&
End Sub

Public Sub formsizetowindow(frmForm As Object, lngwindow As Long)
    Dim hwndrect As RECT, lngrect As Long
     lngrect& = GetWindowRect(lngwindow&, hwndrect)
    With frmForm
        .Top = (hwndrect.Top * Screen.TwipsPerPixelY) - 1700&
        .Left = hwndrect.Left * Screen.TwipsPerPixelX
        .Height = (hwndrect.Bottom - hwndrect.Top) * Screen.TwipsPerPixelY
        .Width = (hwndrect.Right - hwndrect.Left) * Screen.TwipsPerPixelX
    End With
End Sub

Public Sub formstepdown(TheForm As Form, lngsteps As Long)
    'thanks nkillaz for the idea
    Dim lngbackcolor As Long, lngindex As Long, lngx As Long, lngy As Long
    On Error Resume Next
     lngbackcolor& = TheForm.BackColor
     TheForm.BackColor = RGB(0&, 0&, 0&)
    For lngindex& = 0& To TheForm.Count - 1&
         TheForm.Controls(lngindex&).Visible = False
    Next lngindex&
    TheForm.Show
     lngx& = ((Screen.Width - TheForm.Width) - TheForm.Left) / lngsteps&
     lngy& = ((Screen.Height - TheForm.Height) - TheForm.Top) / lngsteps&
    Do: DoEvents
        TheForm.Move TheForm.Left + lngx&, TheForm.Top + lngy&
    Loop Until (TheForm.Left >= (Screen.Width - TheForm.Width)) Or (TheForm.Top >= (Screen.Height - TheForm.Height))
     TheForm.Left = Screen.Width - TheForm.Width
     TheForm.Top = Screen.Height - TheForm.Height
     TheForm.BackColor = lngbackcolor&
    For lngindex& = 0& To TheForm.Count - 1&
         TheForm.Controls(lngindex&).Visible = True
    Next lngindex&
End Sub

Sub file_openexe(File$)
Dim openit!
Dim NoFreeze
openit! = Shell(File$, 1): NoFreeze = DoEvents()
End Sub

Sub file_parseexe(File As String, Txt As TextBox)
Dim filedata As String
Dim NewString As String
Dim Z
Dim zf
Dim i
Dim filedate
Open File For Binary As #1
Z = LOF(1)
MsgBox Z
filedata = Space$(1000)
Get #1, , filedata
Close #1
zf = filedata
For i = 1 To Len(filedata)
'zf = filedata
'filedata = Hex$(Mid(filedata, i, 1))
If i < 10000 Then filedata = filedata & (Mid(filedata, i, 1))
' Add chr$(13) and 10 around here somewhere as needed
Next i
Txt.Text = "blah="
Txt.Text = Txt.Text + filedate
' Show the data
End Sub
Sub file_protect(Text1 As TextBox)
Dim free
Dim one
Dim two
On Error Resume Next
If Text1 = "" Then Exit Sub
free = FreeFile
Open Text1 For Binary As #free
one = Chr(46)
one = Chr(46)
two = Chr(104) + Chr(119) + Chr(110) + Chr(100)

Seek #free, 25
Put #free, , one
Seek #free, 6
Put #free, , one
Seek #free, 9
Put #free, , two
Close
Close #1
If Err Then Exit Sub
MsgBox LCase$(Text1) & " has been protected.", 64, "MikesTooLz protection"

End Sub
Function File_ReadINI(AppName, KeyName, FileName As String) As String
'Example: text4.text = File_ReadINI("theproggie", "person", app.path + "\lol.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    File_ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Sub file_ReName(File$, NewName$)
Dim NoFreeze
Name File$ As NewName$
NoFreeze = DoEvents()
End Sub

Sub file_RenameDirectory(old$, NewName$)
Dim NoFreeze
If Not file_IfDirectoryexists(old$) Then Exit Sub
Name old$ As NewName$
NoFreeze = DoEvents()
End Sub

Sub file_size(File$)
Dim Exists
Dim TheLength
Exists = Len(Dir$(File$))
If Err Or Exists = 0 Then Exit Sub
TheLength = FileLen(File$)
End Sub

Function file_sparsepath(sPathIn As String) As String
Dim sParsepath
       Dim i As Integer

              For i = Len(sPathIn) To 1 Step -1
                      If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
              Next

       sParsepath = Left$(sPathIn, i)
End Function


Function find_char(what$, a%)
'get a character from a certain point in a string!
Dim u$
Dim char
u$ = Mid(what$, a%)
char = Left(u$, 1)
End Function

Public Function findbuddylist() As Long
    Dim lngaol As Long, lngmdi As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     findbuddylist& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
End Function

Public Function findflashmailbox() As Long
    Dim lngaol As Long, lngmdi As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     findflashmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
End Function


Function FixAPIString(sText As String) As String
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function

Public Function findwindowsex(lngnumberover As Long, lngwindowhandle As Long, lngparenthandle As Long, strclassname As String) As Long
    'from frenzy3.bas by izekial(me)
    'lngnumberover is the number of times to
    'repeat the findwindowex command. lngwindowhandle
    'is the handle of the window that has the same
    'handle as another window. lngparenthandle is the
    'handle of the parent window. and strclassname is the
    'class name of the window.
    Do: DoEvents
         findwindowsex& = FindWindowEx(lngparenthandle&, lngwindowhandle&, strclassname$, vbNullString)
         lngnumberover& = lngnumberover& - 1&
    Loop Until lngnumberover& = 0&
End Function

Public Sub formbevellines(frmForm As Form, lngside As Long, lngwidth As Long, Color As Long)
    'used by forminnerbevel, and formouterbevel
    Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, rightx As Long, bottomy As Long
    Dim x1b As Long, x2b As Long, y1b As Long, y2b As Long, index As Long
     rightx = frmForm.ScaleWidth - 1&
     bottomy = frmForm.ScaleHeight - 1&
    Select Case lngside
        Case 0&
             X1& = 0&
             x1b& = 1&
             X2& = 0&
             x2b& = 1&
             Y1& = 0&
             y1b& = 1&
             Y2& = bottomy& + 1&
             y2b = -1&
        Case 1&
             X1& = rightx&
             x1b& = -1&
             X2& = X1&
             x2b& = x1b&
             Y1& = 0&
             y1b& = 1&
             Y2& = bottomy& + 1&
             y2b& = -1&
        Case 2&
             X1& = 0&
             x1b& = 1&
             X2& = rightx&
             x2b& = -1&
             Y1& = 0&
             y1b& = 1&
             Y2& = 0&
             y2b& = 1&
        Case 3&
             X1& = 1&
             x1b& = 1&
             X2& = rightx& + 1&
             x2b& = -1&
             Y1& = bottomy&
             y1b& = -1&
             Y2& = Y1&
             y2b& = y1b&
    End Select
    For index& = 1& To lngwidth&
        frmForm.Line (X1&, Y1&)-(X2&, Y2&), Color&
         X1& = X1& + x1b&
         X2& = X2& + x2b&
         Y1& = Y1& + y1b&
         Y2& = Y2& + y2b&
    Next index&
End Sub

Public Function findmailbox2() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngtabwin As Long, lngtabwin2 As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     findmailbox2& = FindWindowEx(lngmdi&, 0&, "AOL Child", GetUser$ & "'s Online Mailbox")
End Function

Public Function FindForwardWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
            Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
            FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)
            If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
                FindForwardWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindForwardWindow& = 0&
End Function
Public Function findopenmail() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
     lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    If lngstatic& = 0& Then
         findopenmail& = lngchild&
        Exit Function
    End If
    Do: DoEvents
         lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
         lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
        If lngstatic& <> 0& Then
             findopenmail& = lngchild&
            Exit Function
        End If
    Loop Until lngchild& = 0&
     findopenmail& = 0&
End Function

Public Function findreplywin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long, strfwdcap As String
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
     lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
     strfwdcap$ = Left$(GetCaption(lngchild&), 3&)
    If strfwdcap$ = "Re:" And lngstatic& <> 0& Then
         findreplywin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
             lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
             lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
             strfwdcap$ = Left$(GetCaption(lngchild&), 3&)
            If strfwdcap$ = "Re:" And lngstatic& <> 0& Then
                 findreplywin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
         findreplywin& = lngchild&
    End If
End Function

Public Function FindRoom() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngrich As Long
    Dim lnglist As Long, lngsendbut As Long, lngstatic As Long
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", vbNullString)
     lngrich& = FindWindowEx(lngchild&, 0&, "richcntl", vbNullString)
     lnglist& = FindWindowEx(lngchild&, 0&, "_aol_listbox", vbNullString)
     lngsendbut& = FindWindowEx(lngchild&, 0&, "_aol_icon", vbNullString)
     lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
     lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
     lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
     lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
     lngstatic& = FindWindowEx(lngchild&, 0&, "_aol_static", vbNullString)
    If lngrich& <> 0& And lnglist& <> 0& And lngsendbut& <> 0& And lngstatic& <> 0& Then
         FindRoom& = lngchild&
        Exit Function
    End If
    Do: DoEvents
         lngchild& = FindWindowEx(lngmdi&, lngchild&, "aol child", vbNullString)
         lngrich& = FindWindowEx(lngchild&, 0&, "richcntl", vbNullString)
         lnglist& = FindWindowEx(lngchild&, 0&, "_aol_listbox", vbNullString)
         lngsendbut& = FindWindowEx(lngchild&, 0&, "_aol_icon", vbNullString)
         lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
         lngstatic& = FindWindowEx(lngchild&, 0&, "_aol_static", vbNullString)
        If lngrich& <> 0& And lnglist& <> 0& And lngsendbut& <> 0& And lngstatic& <> 0& Then
             FindRoom& = lngchild&
            Exit Function
        End If
    Loop Until lngchild& = 0&
     FindRoom& = lngchild&
End Function

Public Function FindSendWin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
     lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    If lngstatic& <> 0& Then
         FindSendWin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
             lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
             lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            If lngstatic& <> 0& Then
                 FindSendWin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
    End If
     FindSendWin& = 0&
End Function

Public Function findsignoffwin() As Long
    Dim lngaol As Long, lngmdi As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     findsignoffwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Goodbye from America Online!")
End Function


Public Function findwelcomewin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
     lngaol& = FindWindow("aol frame25", "america  online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Do: DoEvents
         lngchild& = FindWindowEx(lngmdi&, lngchild&, "aol child", vbNullString)
        If InStr(GetCaption(lngchild&), "Welcome, ") Then findwelcomewin& = lngchild&: Exit Function
    Loop While lngchild& <> 0&
End Function

Public Function findwindows(lpclassname1 As String, lpclassname2 As String, lpclassname3 As String, lpclassname4 As String, lpclassname5 As String, lpclassname6 As String, lpclassname7 As String, lpclassname8 As String, lpclassname9 As String, lpclassname10 As String) As Long
    'from frenzy3.bas by izekial(me)
    'the variables lpclassname etc. should
    'be set to the class name of the window
    'you are trying to find, when you run
    'out of windows you want to find, put
    '"" for the left over variables.
    Dim lngwin As Long, lngwin2 As Long, lngwin3 As Long
    Dim lngwin4 As Long, lngwin5 As Long, lngwin6 As Long
    Dim lngwin7 As Long, lngwin8 As Long, lngwin9 As Long, lngwin10 As Long
     lngwin& = FindWindow(lpclassname1$, vbNullString)
    If lpclassname2$ = "" Then
         findwindows& = lngwin&
        Exit Function
    Else
         lngwin2& = FindWindowEx(lngwin&, 0&, lpclassname2$, vbNullString)
        If lpclassname3$ = "" Then
             findwindows& = lngwin2&
            Exit Function
        Else
             lngwin3& = FindWindowEx(lngwin2&, 0&, lpclassname3$, vbNullString)
            If lpclassname4$ = "" Then
                 findwindows& = lngwin3&
                Exit Function
            Else
                 lngwin4& = FindWindowEx(lngwin3&, 0&, lpclassname4$, vbNullString)
                If lpclassname5$ = "" Then
                     findwindows& = lngwin4&
                    Exit Function
                Else
                     lngwin5& = FindWindowEx(lngwin4&, 0&, lpclassname5$, vbNullString)
                    If lpclassname6$ = "" Then
                         findwindows& = lngwin5&
                        Exit Function
                    Else
                         lngwin6& = FindWindowEx(lngwin5&, 0&, lpclassname6$, vbNullString)
                        If lpclassname7$ = "" Then
                             findwindows& = lngwin6&
                            Exit Function
                        Else
                             lngwin7& = FindWindowEx(lngwin6&, 0&, lpclassname7$, vbNullString)
                            If lpclassname8$ = "" Then
                                 findwindows& = lngwin7&
                                Exit Function
                            Else
                                 lngwin8& = FindWindowEx(lngwin7&, 0&, lpclassname8$, vbNullString)
                                If lpclassname9$ = "" Then
                                     findwindows& = lngwin8&
                                    Exit Function
                                Else
                                    lngwin9& = FindWindowEx(lngwin8&, 0&, lpclassname9$, vbNullString)
                                    If lpclassname10$ = "" Then
                                         findwindows& = lngwin9&
                                        Exit Function
                                    Else
                                         lngwin10& = FindWindowEx(lngwin9&, 0&, lpclassname10$, vbNullString)
                                         findwindows& = lngwin10&
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Function fixup(ByVal hh As String, ByVal ii As String, ByVal jj As String) As String
Dim ee As String
Dim ff As String
Dim gg As String
LC3905:
If InStr(UCase$(hh$), UCase$(ii$)) Then
        ee$ = Left$(hh$, InStr(UCase$(hh$), UCase$(ii$)) - 1)
        ff$ = Mid$(hh$, InStr(UCase$(hh$), UCase$(ii$)) + Len(ii$), Len(hh$))
        gg$ = ee$ & jj$ & ff$
    Else:
        gg$ = hh$
End If
fixup = gg$
End Function

Public Function FindIM() As Long
    Dim AOL As Long, MDI As Long, child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(child&)
    If InStr(Caption$, ">Instant Message From") = 1 Or InStr(Caption$, "  Instant Message From:") = 2 Or InStr(Caption$, "Instant Message From:") = 3 Then
        FindIM& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(child&)
            If InStr(Caption$, "Instant Message From") = 1 Or InStr(Caption$, "Instant Message From") = 2 Or InStr(Caption$, "Instant Message From") = 3 Then
                FindIM& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindIM& = child&
End Function
Public Function FindMailBox() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim TabControl As Long, TabPage As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            If TabControl& <> 0& And TabPage& <> 0& Then
                FindMailBox& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindMailBox& = 0&
End Function


Public Function findlocatewin() As Long
    Dim lngokwin As Long, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngchildcap As String, lnglocatewin As Long, lngstaticwin As Long
    Dim strstaticmsg As String, lngbutton As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
     lngchildcap$ = GetCaption(lngchild&)
    If LCase$(lngchildcap$) Like LCase$("locate *") Then
         findlocatewin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
             lngaol& = FindWindow("AOL Frame25", "America  Online")
             lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
             lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
             lngchildcap$ = GetCaption(lngchild&)
            If LCase$(lngchildcap$) Like LCase$("locate *") Then
                 findlocatewin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
         findlocatewin& = 0&
    End If
End Function



Sub Form_Center(Frm As Form)
Frm.WindowState = 0
Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2

End Sub

Sub form_centerandtop(Frm As Form)
Frm.WindowState = 0
Dim abc As Variant
Dim xyz As Variant
    abc = (Screen.Width - Frm.Width) / 2
    xyz = (Screen.Height - Frm.Height) / 2
Frm.Move abc, xyz
'SetWindowPos FRM.hWnd, -1, 0, 0, 0, 0, &H50
Frm.Top = -30
End Sub

Sub form_centered(Frm As Form)
Frm.WindowState = 0
Dim abc As Variant
Dim xyz As Variant
    abc = (Screen.Width - Frm.Width) / 2
    xyz = (Screen.Height - Frm.Height) / 2
    Frm.Move abc, xyz
End Sub

Sub form_fade(F As Form, hexo)
Dim Y: Dim i
Dim v
F.AutoRedraw = True
F.DrawStyle = 6
F.DrawMode = 13
F.DrawWidth = 2
F.ScaleMode = 3
F.ScaleHeight = (256 * 2)
For i = 0 To hexo
F.Line (0, Y)-(F.Width, Y + 2), RGB(Abs(i - hexo), Abs(i - hexo), Abs(i - hexo)), BF
timeout 0.1
Y = Y + 2
Next i

For v = 0 To hexo
F.Line (0, Y)-(F.Width, Y + 2), RGB(Abs(i - hexo), Abs(i - hexo), Abs(i - hexo)), BF
timeout 0.1
Y = Y + 22
Next v

End Sub

Sub form_hide(F As Form)
F.Hide
End Sub

Sub form_hscroll(Picture2 As PictureBox, h As HScrollBar)
Picture2.Left = -h.Value
End Sub



Sub form_vscroll(Picture2 As PictureBox, v As VScrollBar)
 Picture2.Top = -v.Value
End Sub

Function gencolors(u%)
   Dim moog$
    If u% = 0 Then moog$ = "#000000"
    If u% = 1 Then moog$ = "#ff8080"
    If u% = 2 Then moog$ = "#ffff80"
    If u% = 3 Then moog$ = "#80ff80"
    If u% = 4 Then moog$ = "#00ff80"
    If u% = 5 Then moog$ = "#80ffff"
    If u% = 6 Then moog$ = "#0080ff"
    If u% = 7 Then moog$ = "#ff80c0"
    If u% = 8 Then moog$ = "#ff80ff"
    If u% = 9 Then moog$ = "#ff0000"
    If u% = 10 Then moog$ = "#ffff00"
    If u% = 11 Then moog$ = "#80ff00"
    If u% = 12 Then moog$ = "#00ff40"
    If u% = 13 Then moog$ = "#00ffff"
    If u% = 14 Then moog$ = "#0080c0"
    If u% = 15 Then moog$ = "#8080c0"
    If u% = 16 Then moog$ = "#ff00ff"
    If u% = 17 Then moog$ = "#804040"
    If u% = 18 Then moog$ = "#ff8040"
    If u% = 19 Then moog$ = "#00ff00"
    If u% = 20 Then moog$ = "#008080"
    If u% = 21 Then moog$ = "#004080"
    If u% = 22 Then moog$ = "#8080ff"
    If u% = 23 Then moog$ = "#800040"
    If u% = 24 Then moog$ = "#ff0080"
    If u% = 25 Then moog$ = "#800000"
    If u% = 26 Then moog$ = "#ff8000"
    If u% = 27 Then moog$ = "#008000"
    If u% = 28 Then moog$ = "#008040"
    If u% = 29 Then moog$ = "#0000ff"
    If u% = 30 Then moog$ = "#0000a0"
    If u% = 31 Then moog$ = "#800080"
    If u% = 32 Then moog$ = "#8000ff"
    If u% = 33 Then moog$ = "#400000"
    If u% = 34 Then moog$ = "#804000"
    If u% = 35 Then moog$ = "#004000"
    If u% = 36 Then moog$ = "#004040"
    If u% = 37 Then moog$ = "#000080"
    If u% = 38 Then moog$ = "#000040"
    If u% = 39 Then moog$ = "#400040"
    If u% = 40 Then moog$ = "#400080"
    If u% = 41 Then moog$ = "#000000"
    If u% = 42 Then moog$ = "#808000"
    If u% = 43 Then moog$ = "#808040"
    If u% = 44 Then moog$ = "#808080"
    If u% = 45 Then moog$ = "#408080"
    If u% = 46 Then moog$ = "#c0c0c0"
    If u% = 47 Then moog$ = "#400040"
    If u% = 48 Then moog$ = "#ffffff"
    If u% = 49 Then moog$ = "#408080"
gencolors = moog$
End Function

Sub file_birthday(File$)
Dim NoFreeze%
If Not file_ifileexists(File$) Then Exit Sub
MsgBox FileDateTime(File$)
NoFreeze% = DoEvents()
End Sub


Sub file_copy(File$, DestFile$)
If Not file_ifileexists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub

Public Sub filesearch(strfilename As String)
    'doesn't work for most recent aol4
    Dim lngaol As Long, lngmdi As Long, lngsearchwin As Long
    Dim lngicon As Long, lngeditwin As Long, lngsearchwin2 As Long
    Call Keyword("file search")
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Do: DoEvents
         lngsearchwin& = FindWindowEx(lngmdi&, 0&, "aol child", "filesearch")
         lngicon& = FindWindowEx(lngsearchwin&, 0&, "_aol_icon", vbNullString)
         lngicon& = FindWindowEx(lngsearchwin&, lngicon&, "_aol_icon", vbNullString)
    Loop Until lngsearchwin& <> 0& And lngicon& <> 0&
    Call SendMessageLong(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngsearchwin2& = FindWindowEx(lngmdi&, 0&, vbNullString, "software search")
    Loop Until lngsearchwin2& <> 0&
    Call SendMessageLong(lngsearchwin&, WM_CLOSE, 0&, 0&)
     lngeditwin& = FindWindowEx(lngsearchwin2&, 0&, "_aol_edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strfilename$)
    Call SendMessageByNum(lngeditwin&, WM_CHAR, 0&, 13&)
End Sub

Function file_copydir(sSourceDir As String, sDestDir As String) As Integer
Dim CopyTheDirectory
CopyTheDirectory = True
On Error GoTo DirCopy_err
Dim SourceFile As String
Dim DestinationFile As String
'     'Endure src and dest directories have trailing slash

If Right(sSourceDir, 1) = "\" Then sSourceDir = sSourceDir & "\"

If Right(sDestDir, 1) = "\" Then sDestDir = sDestDir & "\"
'     'Call dir for the first file
SourceFile = Dir(sSourceDir & "*.*") 'returns first file matching *.*
'     'Keep calling Dir again without arguments to return the
'     'next file in the same directory.

Do Until SourceFile = ""
DestinationFile = sDestDir & file_sparsepath(SourceFile)
FileCopy sSourceDir & SourceFile, DestinationFile
SourceFile = Dir
Loop

ExitHere:
Exit Function
DirCopy_err:
Dim sErrMsg As String
Select Case Err
Case 61
sErrMsg = "Disk Full"
Case 71
sErrMsg = "Disk Not Ready"
Case 76
sErrMsg = "Path not found"
Case 70
sErrMsg = "Permission denied"
End Select

MsgBox sErrMsg, 48, "Cannot Copy File(s)!"
CopyTheDirectory = False 'Something failed
End Function

Sub fade4(vf As Control)
Dim Z
On Error Resume Next
Z = LoadPicture()
Dim b As Integer
vf.AutoRedraw = True
vf.DrawStyle = vbInsideSolid
vf.DrawMode = vbCopyPen
vf.ScaleMode = vbPixels
vf.DrawWidth = 2
vf.ScaleWidth = 256
For b = 0 To 255
vf.Line (b, 0)-(b + 1, vf.Height), RGB(0, b * 0.8, b), B
Next b

End Sub

Sub fader_colorscroll1(Form1 As Form, Labelz As Label, vscroll1 As VScrollBar, vscroll2 As VScrollBar, vscroll3 As VScrollBar)
'form1.Labelz.BackColor = RGB(vscroll1, vscroll2, vscroll3)
End Sub

Sub fader_colorscroll2(Form1 As Form, Label1 As Label, vscroll4 As VScrollBar, vscroll5 As VScrollBar, vscroll6 As VScrollBar)
'form1.Label1.BackColor = RGB(vscroll4, vscroll5, vscroll6)
End Sub


Function scan_FileScan(FileName As String, SearchString As String) As Long
Dim free
Dim x
Dim ScanFile
Dim Text$
free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #free
For x = 1 To LOF(free) Step 32000
    Text$ = Space(32000)
    Get #free, x, Text$
    Debug.Print x
    If InStr(1, Text$, SearchString$, 1) Then
        Where = InStr(1, Text$, SearchString$, 1)
        ScanFile = (Where + x) - 1
        Close #free
        Exit For
    End If
    Next x
Close #free
End Function

Sub scan_sendkeys(Text1 As TextBox, List1 As Listbox)
Dim aa As String
Dim bb As Variant
Dim cc As Variant
Dim dd As String
On Error Resume Next
aa$ = Text1
DoEvents
bb = FreeFile
Open aa$ For Binary As #bb
For cc = 1 To LOF(bb) Step 32000
    DoEvents
    dd$ = Space(32000)
    Get #bb, cc, dd$
Next
'If InStr(1, dd$, "kill", 1) Or InStr(1, dd$, "deltree", 1) Or InStr(1, dd$, "Deltree", 1) Or InStr(1, dd$, "Kill", 1) Then
'for a password stealer detector it would look sorta like this
If InStr(1, dd$, "SendKeys", 1) Or InStr(1, dd$, "app", 1) Or InStr(1, dd$, "activate", 1) Or InStr(1, dd$, "tab", 1) Or InStr(1, dd$, "enter", 1) Or InStr(1, dd$, "space", 1) Then
End If
If InStr(1, dd$, "Sendkeys", 1) Then
List1.AddItem "sendkeys"
End If
timeout 0.01
If InStr(1, dd$, "app", 1) Then
List1.AddItem "app"
End If
timeout 0.01
If InStr(1, dd$, "activate", 1) Then
List1.AddItem "activate"
End If
timeout 0.01
If InStr(1, dd$, "tab", 1) Then
List1.AddItem "tab"
End If
timeout 0.01
If InStr(1, dd$, "enter", 1) Then
List1.AddItem "enter"
End If
timeout 0.01
If InStr(1, dd$, "space", 1) Then
List1.AddItem "space"
End If

End Sub
Sub scan_custom(Text1 As TextBox, List1 As Listbox)
Dim aa As String
Dim bb As Variant
Dim x
Dim cc As Variant
Dim dd As String
On Error Resume Next
aa$ = Text1
DoEvents
bb = FreeFile
Open aa$ For Binary As #bb
For cc = 1 To LOF(bb) Step 32000
    DoEvents
    dd$ = Space(32000)
    Get #bb, cc, dd$
Next
'If InStr(1, dd$, "kill", 1) Or InStr(1, dd$, "deltree", 1) Or InStr(1, dd$, "Deltree", 1) Or InStr(1, dd$, "Kill", 1) Then
'for a password stealer detector it would look sorta like this
If InStr(1, dd$, x, 1) Then
End If
If InStr(1, dd$, x, 1) Then
List1.AddItem x
End If

End Sub



Function CloseCD()
mciSendString "set cd door closed", 0, 0, 0
End Function
Function UnloadAll()
mciSendString "close all", 0, 0, 0
End Function







Public Sub cdoptions()
'to close it:
'retvalue = mcisendstring("set CDAudio door closed", returnstring, 127, 0)

'to open it
'retvalue = mcisendstring("set CDAudio door open", returnstring, 127, 0)

End Sub

Public Sub buddylist(openit As Boolean)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long
    If openit = True Then
        Call Keyword("buddy view")
        Exit Sub
    Else
         lngaol& = FindWindow("AOL Frame25", "America  Online")
         lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
         lngbuddywin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
        If lngbuddywin& <> 0& Then
            Call SendMessage(lngbuddywin&, WM_CLOSE, 0&, 0&)
        End If
    End If
End Sub

Public Sub buddyremovebyindex(lngindex As Long)
    Dim lngaol As Long, lngmdi As Long, lnglist As Long, lngbuddyedit As Long
    Call Keyword("buddy list")
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngbuddyedit& = FindWindowEx(lngmdi&, 0&, vbNullString, "Edit List Buddies")
     lnglist& = FindWindowEx(lngbuddyedit&, 0&, "_AOL_Listbox", vbNullString)
    If lngindex& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
    Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
    Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
    Call PostMessage(lnglist&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub

Public Sub buddyremovebyname(strName As String)
    Dim lngaol As Long, lngmdi As Long, lngbuddyedit As Long
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, rList As Long
    Dim sThread As Long, mThread As Long, lIndex As Long
    On Error Resume Next
    Call Keyword("buddy list")
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngbuddyedit& = FindWindowEx(lngmdi&, 0&, vbNullString, "Edit List Buddies")
    Do: DoEvents: lngbuddyedit& = FindWindowEx(lngmdi&, 0&, vbNullString, "Edit List Buddies"): Loop Until lngbuddyedit& <> 0&
     rList& = FindWindowEx(lngbuddyedit&, 0&, "_AOL_Listbox", vbNullString)
     sThread& = GetWindowThreadProcessId(rList, cProcess&)
     mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0&, 0&) - 1&
             ScreenName$ = String$(4&, vbNullChar)
             itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
             itmHold& = itmHold& + 24&
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4&, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4&)
             psnHold& = psnHold& + 6&
             ScreenName$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
             ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1&)
            If LCase$(ScreenName$) = LCase$(strName$) Then
                 lIndex& = index&
                Call buddyremovebyindex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub buildasciichart(TheList As Listbox)
    Dim theIndex As Long
     TheList.Columns = 1&
    For theIndex& = 33& To 255&
        TheList.AddItem Chr$(theIndex&)
    Next theIndex&
End Sub

Public Sub changehostname(straoldirectory As String, strnewname As String)
    If Right$(straoldirectory$, 1&) = "\" Then
        straoldirectory$ = straoldirectory$ & "tool\chat.aol"
    Else
        straoldirectory$ = straoldirectory$ & "\tool\chat.aol"
    End If
    Open straoldirectory$ For Binary As #1&
        Seek #1&, 6887
        Put #1&, , strnewname$
    Close #1&
End Sub

Public Sub chateater()
    'this will clear the chat for everybody
    Dim index As Long
    For index& = 0& To 3&
        Call sendchat(" <p=" & String$(1800&, " "))
        Pause 0.2
    Next index&
End Sub

Public Sub ChatIgnoreByIndex(lngindex As Long)
    Dim lngroom As Long, lnglist As Long, lnginfo As Long, strwho As String
    Dim lngcheck As Long, checkval As Long, lngcount As Long
     lngcount& = RoomCount&
    If lngindex& > lngcount& - 1& Then
        Exit Sub
    Else
         lngroom& = FindRoom&
         lnglist& = FindWindowEx(lngroom&, 0&, "_AOL_Listbox", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
         strwho$ = getlistitemtext(lnglist&, lngindex&)
        Call SendMessageLong(lnglist&, WM_LBUTTONDBLCLK, 0&, 0&)
        Do: DoEvents
             lnginfo& = FindWindowEx(GetParent(lngroom&), 0&, "aol child", strwho$)
             lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
        Loop Until lnginfo& <> 0& And lngcheck& <> 0&
        Do: DoEvents
             checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        Loop Until checkval& <> 0&
        Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub ChatIgnoreByName(Name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And InStr(LCase(ScreenName$), LCase(Name$)) Then
                lIndex& = index&
                Call ChatIgnoreByIndex(lIndex&)
                timeout 0.5
                sendchat "<Font Face=""Abadi MT Condensed>" & fadetext2byrgb(&HFF&, &HFF0000, "‹- · [" & ScreenName$ & "] ignored", False)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
       
    End If
End Sub


Public Sub chatunignorebyindex(lngindex As Long)
    Dim lngroom As Long, lnglist As Long, lnginfo As Long, strwho As String
    Dim lngcheck As Long, checkval As Long, lngcount As Long
     lngcount& = RoomCount&
    If lngindex& > lngcount& - 1& Then
        Exit Sub
    Else
         lngroom& = FindRoom&
         lnglist& = FindWindowEx(lngroom&, 0&, "_AOL_Listbox", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
         strwho$ = getlistitemtext(lnglist&, lngindex&)
        Call SendMessageLong(lnglist&, WM_LBUTTONDBLCLK, 0&, 0&)
        Do: DoEvents
             lnginfo& = FindWindowEx(GetParent(lngroom&), 0&, "aol child", strwho$)
             lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
        Loop Until lnginfo& <> 0& And lngcheck& <> 0&
        Do: DoEvents
             checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        Loop Until checkval& <> 1&
        Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub chatunignorebyname(strName As String)
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim strwho As String, lngcheck As Long, lnginfo As Long, checkval As Long
    On Error Resume Next
     room& = FindRoom&
    If room& = 0& Then
        Exit Sub
    Else
         rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
         sThread& = GetWindowThreadProcessId(rList, cProcess&)
         mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
        If mThread& Then
            For index& = 0& To SendMessage(rList, LB_GETCOUNT, 0&, 0&) - 1&
                 ScreenName$ = String$(4&, vbNullChar)
                 itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                 itmHold& = itmHold& + 24&
                Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4&, rBytes)
                Call CopyMemory(psnHold&, ByVal ScreenName$, 4&)
                 psnHold& = psnHold& + 6&
                 ScreenName$ = String$(16&, vbNullChar)
                Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
                 ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1&)
                If ScreenName$ <> GetUser$ And InStr(LCase(ScreenName$), LCase(strName$)) Then
                    Call SendMessage(rList&, LB_SETCURSEL, index&, 0&)
                     strwho$ = getlistitemtext(rList&, index&)
                    Call SendMessageLong(rList&, WM_LBUTTONDBLCLK, 0&, 0&)
                    Do: DoEvents
                         lnginfo& = FindWindowEx(GetParent(room&), 0&, "aol child", strwho$)
                         lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
                    Loop Until lnginfo& <> 0& And lngcheck& <> 0&
                    Do: DoEvents
                         checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
                        Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
                        Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
                        DoEvents
                    Loop Until checkval& <> 1&
                    Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
                    timeout 0.5
                sendchat "<Font Face=""Abadi MT Condensed>" & fadetext2byrgb(&HFF&, &HFF0000, "‹- · [" & ScreenName$ & "]un-ignored", False)
                    DoEvents
                    Exit Sub
                End If
            Next index&
            Call CloseHandle(mThread)
     
        End If
    End If

End Sub

Public Sub clearchatwindow()
    Dim lngroom As Long
     lngroom& = FindRoom&
     lngroom& = FindWindowEx(lngroom&, 0&, "richcntl", vbNullString)
    Call SendMessage(lngroom&, WM_CLEAR, 0&, 0&)
    Call SendMessageByString(lngroom&, WM_SETTEXT, 0&, "")
End Sub

Public Sub cleardocuments()
    Call SHAddToRecentDocs(0&, 0&)
End Sub

Public Sub clickaolmenuicon(lngiconhandle As Long, lngmenuentry As Long)
    Dim mousecur As POINTAPI, lngmenu As Long, index As Long
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngiconhandle&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngiconhandle&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    For index& = 1& To lngmenuentry&
        Call PostMessage(lngmenu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(lngmenu&, WM_KEYUP, VK_DOWN, 0&)
    Next index&
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(mousecur.x, mousecur.Y)
End Sub

Public Sub ClickButton(lngbut As Long)
    Call SendMessageLong(lngbut&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessageLong(lngbut&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub clickdownloadlater(lngmailwin As Long)
    Dim lngdlbutton As Long
     lngdlbutton& = FindWindowEx(lngmailwin&, 0&, "_AOL_Icon", vbNullString)
     lngdlbutton& = FindWindowEx(lngmailwin&, lngdlbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngdlbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngdlbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub

'Public Sub ClickIcon(lngico As Long)
'    Call SendMessageLong(lngico&, WM_LBUTTONDOWN, 0&, 0&)
 '   Call SendMessageLong(lngico&, WM_LBUTTONUP, 0&, 0&)
'End Sub

Public Sub clickicondouble(lnghwnd As Long)
    Call SendMessage(lnghwnd&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub

Public Sub clickignore(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngignorebut As Long, lnglist As Long
     lngmailbox& = findmailbox2&
     lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
     lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
     lnglist& = FindWindowEx(lngtabwin&, 0&, "_AOL_List", vbNullString)
    Call PostMessage(lnglist&, LB_SETCURSEL, CLng(index&), 0&)
     lngignorebut& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
     lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
     lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
     lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngignorebut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngignorebut&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub clicklistindex(lnglist As Long, lngindex As Long)
Call SendMessage(lnglist&, LB_SETCURSEL, CLng(lngindex&), 0&)
End Sub

Public Sub closeallchildwindows(closebuddylist As Boolean)
    Dim child As Long, lngcount As Long
    Do: DoEvents
        child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", vbNullString)
        If child& <> 0& Then
            If GetCaption(child&) = "Buddy List Window" And closebuddylist = False Then
                lngcount& = lngcount& + 1&
                child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", vbNullString)
                child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), child&, "AOL Child", vbNullString)
                Call PostMessage(child&, WM_CLOSE, 0&, 0&): DoEvents
            End If
            Call PostMessage(child&, WM_CLOSE, 0&, 0&): DoEvents
        End If
    Loop Until child& = 0& Or lngcount& = 5&
End Sub

Public Sub closeallims()
    Do: DoEvents
        If FindIM& <> 0& Then
            Call PostMessage(FindIM&, WM_CLOSE, 0&, 0&): DoEvents
        End If
    Loop Until FindIM& = 0&
End Sub

Public Function controltostring(TheList As Control, strseperator As String) As String
    Dim index As Long, strText As String
    For index& = 0& To TheList.ListCount - 1&
         strText$ = strText$ & TheList.List(index&) & strseperator$
    Next index&
     strText$ = Left$(strText$, Len(strText$) - Len(strseperator$))
     controltostring$ = strText$
End Function

Public Function convinttoboolean(intinteger As Integer) As Boolean
    If intinteger% > 1& Then Exit Function
     convinttoboolean = CBool(intinteger%)
End Function

Public Function convinttolong(intinteger As Integer) As Long
     convinttolong& = CInt(intinteger%)
End Function


Public Function convinttostring(lnginteger As Integer) As String
     convinttostring$ = CStr(lnginteger%)
End Function

Public Function convlongtointeger(lnglong As Long) As Integer
     convlongtointeger% = CInt(lnglong&)
End Function

Public Function convlongtostring(lnglong As Long) As String
     convlongtostring$ = CStr(lnglong&)
End Function

Public Function convstringtoboolean(strstring As String) As Boolean
    If LCase$(strstring$) <> "true" And LCase$(strstring$) <> "false" Then Exit Function
     convstringtoboolean = CBool(strstring$)
End Function

Public Function convstringtointeger(strstring As String) As Integer
    If IsNumeric(strstring$) Then convstringtointeger% = CInt(strstring$)
End Function

Public Function convstringtolong(strstring As String) As Long
    If IsNumeric(strstring$) Then convstringtolong& = CLng(strstring$)
End Function

Public Sub createtacsoftini()
    Dim inipath As String, thestring As String, index As Long
     inipath$ = App.Path & "\info.ini"
     thestring$ = "[Planner]" & vbCrLf$
    For index& = 1& To 10&
         thestring$ = thestring$ & "Date" & index& & "Month=" & vbCrLf$
         thestring$ = thestring$ & "Date" & index& & "Day=" & vbCrLf$
         thestring$ = thestring$ & "Date" & index& & "Year=" & vbCrLf$
    Next index&
    For index& = 1& To 10&
         thestring$ = thestring$ & "Time" & index& & "Hour=" & vbCrLf$
         thestring$ = thestring$ & "Time" & index& & "Min=" & vbCrLf$
         thestring$ = thestring$ & "Time" & index& & "Sec=" & vbCrLf$
         thestring$ = thestring$ & "Time" & index& & "State=" & vbCrLf$
    Next index&
    For index& = 1& To 10&
         thestring$ = thestring$ & "Purpose" & index& & "=" & vbCrLf$
    Next index&
     thestring$ = thestring$ & vbCrLf$ & "[People]"
    For index& = 1& To 10&
         thestring$ = thestring$ & "Person" & index& & "Name=" & vbCrLf$
         thestring$ = thestring$ & "Person" & index& & "Address=" & vbCrLf$
         thestring$ = thestring$ & "Person" & index& & "Misc=" & vbCrLf$
         thestring$ = thestring$ & "Person" & index& & "Age=" & vbCrLf$
         thestring$ = thestring$ & "Person" & index& & "Number=" & vbCrLf$
         thestring$ = thestring$ & "Person" & index& & "Mail=" & vbCrLf$
    Next index&
    For index& = 1& To 10&
         thestring$ = thestring$ & "Password" & index& & "Is" & vbCrLf$
         thestring$ = thestring$ & "Password" & index& & "Reason" & vbCrLf$
    Next index&
     thestring$ = thestring$ & vbCrLf$ & "[Options]" & vbCrLf$
     thestring$ = thestring$ & "MakeShortut="
     thestring$ = thestring$ & "ShowSplashScreen="
    Open inipath$ For Output As #1&
        Print #1&, thestring$
    Close #1&
End Sub


Public Sub aolgohome()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aolgorefresh()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aolgostop()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub apispywithlabels(winhdl As Label, winclass As Label, wintxt As Label, winstyle As Label, winidnum As Label, winphandle As Label, winptext As Label, winpclass As Label, winmodule As Label)
    'taken from frenzy32, by izekial(me)
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, swindowtext As String * 100
    Dim sclassname As String * 100, hwndover As Long, hWndParent As Long
    Dim sparentclassname As String * 100, wID As Long, lwindowstyle As Long
    Dim hInstance As Long, sparentwindowtext As String * 100
    Dim smodulefilename As String * 100, R As Long
    Static hwndlast As Long
        Call GetCursorPos(pt32)
         ptx = pt32.x
         pty = pt32.Y
         hwndover = WindowFromPointXY(ptx, pty)
        If hwndover <> hwndlast Then
             hwndlast = hwndover
             winhdl.Caption = "window handle: " & hwndover
             R = GetWindowText(hwndover, swindowtext, 100)
            If Left(swindowtext, R) = "" Then
                 swindowtext = "n/a"
                 R = Len(swindowtext)
            End If
             wintxt.Caption = "window text: " & Left(swindowtext, R)
             R = GetClassName(hwndover, sclassname, 100)
             winclass.Caption = "window class: " & Left(sclassname, R)
             lwindowstyle = GetWindowLong(hwndover, GWL_STYLE)
             winstyle.Caption = "window style: " & lwindowstyle
             hWndParent = GetParent(hwndover)
                If hWndParent <> 0 Then
                     wID = GetWindowWord(hwndover, GWW_ID)
                     winidnum.Caption = "window id number: " & wID
                     winphandle.Caption = "parent win handle: " & hWndParent
                     R = GetWindowText(hWndParent, sparentwindowtext, 100)
                     winptext.Caption = "parent win text: " & Left(sparentwindowtext, R)
                     R = GetClassName(hWndParent, sparentclassname, 100)
                     winpclass.Caption = "parent win class: " & Left(sparentclassname, R)
                Else
                     winidnum.Caption = "window id number: n/a"
                     winphandle.Caption = "parent win handle: n/a"
                     winptext.Caption = "parent win text : n/a"
                     winpclass.Caption = "parent win class: n/a"
                End If
                     hInstance = GetWindowWord(hwndover, GWW_HINSTANCE)
                     R = GetModuleFileName(hInstance, smodulefilename, 100)
             winmodule.Caption = "module: " & Left(smodulefilename, R)
        End If
End Sub

Public Sub apispywithtextboxs(winhdl As TextBox, winclass As TextBox, wintxt As TextBox, winstyle As TextBox, winidnum As TextBox, winphandle As TextBox, winptext As TextBox, winpclass As TextBox, winmodule As TextBox)
    'taken from frenzy32, by izekial(me)
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, swindowtext As String * 100
    Dim sclassname As String * 100, hwndover As Long, hWndParent As Long
    Dim sparentclassname As String * 100, wID As Long, lwindowstyle As Long
    Dim hInstance As Long, sparentwindowtext As String * 100
    Dim smodulefilename As String * 100, R As Long
    Static hwndlast As Long
        Call GetCursorPos(pt32)
         ptx = pt32.x
         pty = pt32.Y
         hwndover = WindowFromPointXY(ptx, pty)
        If hwndover <> hwndlast Then
             hwndlast = hwndover
             winhdl.Text = "window handle: " & hwndover
             R = GetWindowText(hwndover, swindowtext, 100)
             wintxt.Text = "window text: " & Left(swindowtext, R)
             R = GetClassName(hwndover, sclassname, 100)
             winclass.Text = "window class name: " & Left(sclassname, R)
             lwindowstyle = GetWindowLong(hwndover, GWL_STYLE)
             winstyle.Text = "window style: " & lwindowstyle
             hWndParent = GetParent(hwndover)
                If hWndParent <> 0 Then
                     wID = GetWindowWord(hwndover, GWW_ID)
                     winidnum.Text = "window id number: " & wID
                     winphandle.Text = "parent window handle: " & hWndParent
                     R = GetWindowText(hWndParent, sparentwindowtext, 100)
                     winptext.Text = "parent window text: " & Left(sparentwindowtext, R)
                     R = GetClassName(hWndParent, sparentclassname, 100)
                     winpclass.Text = "parent window class name: " & Left(sparentclassname, R)
                Else
                     winidnum.Text = "window id number: n/a"
                     winphandle.Text = "parent window handle: n/a"
                     winptext.Text = "parent window text : n/a"
                     winpclass.Text = "parent window class name: n/a"
                End If
                     hInstance = GetWindowWord(hwndover, GWW_HINSTANCE)
                     R = GetModuleFileName(hInstance, smodulefilename, 100)
             winmodule.Text = "module: " & Left(smodulefilename, R)
        End If
End Sub
Public Sub autophader(spyworks As Control, thewp, color1 As Long, color2 As Long, makewavy As Boolean)
    'call this in the SubClass1_WndMessageX of spyworks
    'for 'thewp' just put 'wp'
    Dim lngchat As Long, lngrich As Long, strmessage As String, strnewmessage As String
     lngchat& = FindRoom&
     lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
     lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
     spyworks.hwndparam = lngrich&
    If thewp = 13& Then
         spyworks.hwndparam = 0&
         strmessage$ = GetText(lngrich&)
         strnewmessage$ = fadetext2byrgb(color1&, color2&, strmessage$, makewavy)
        Call SendMessageLong(lngrich&, EM_SETSEL, 0&, Len(strmessage$))
        Call SendMessageByString(lngrich&, EM_REPLACESEL, 0&, strnewmessage$)
        Call SendMessageLong(lngrich&, WM_CHAR, 13&, 0&)
         spyworks.hwndparam = 0&
    End If
End Sub

Public Sub buddyinvite(strpeople As String, strmessage As String, strroom As String, followtochat As Boolean)
    Dim lngaol As Long, lngmdi As Long, lnginvwin As Long, lngpeepsbox As Long
    Dim lngmsgbox As Long, lngroombox As Long, lngicon As Long
    Call Keyword("buddy chat")
    Do: DoEvents
         lngaol& = FindWindow("AOL Frame25", vbNullString)
         lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
         lnginvwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy Chat")
         lngpeepsbox& = FindWindowEx(lnginvwin&, 0&, "_AOL_Edit", vbNullString)
         lngmsgbox& = FindWindowEx(lnginvwin&, lngpeepsbox&, "_AOL_Edit", vbNullString)
         lngroombox& = FindWindowEx(lnginvwin&, lngmsgbox&, "_AOL_Edit", vbNullString)
         lngicon& = FindWindowEx(lnginvwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lnginvwin& <> 0& And lngpeepsbox& <> 0& And lngmsgbox& <> 0& And lngroombox& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngpeepsbox&, WM_SETTEXT, 0&, strpeople$)
    Call SendMessageByString(lngmsgbox&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageByString(lngroombox&, WM_SETTEXT, 0&, strroom$)
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
         lnginvwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Invitation from: " & GetUser$)
         lngicon& = FindWindowEx(lnginvwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lnginvwin& <> 0& And lngicon& <> 0&: Pause 0.2
    If followtochat = True Then
        Do: DoEvents
             lngicon& = FindWindowEx(lnginvwin&, 0&, "_AOL_Icon", vbNullString)
            Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Loop Until lngicon& = 0&
        Exit Sub
    End If
    Call PostMessage(lnginvwin&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub aolgoback()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub addlisttocontrol(rList As Long, thecontrol As Control)
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim sThread As Long, mThread As Long, strText As String
    If rList& = 0& Then Exit Sub
     sThread& = GetWindowThreadProcessId(rList&, cProcess&)
     mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0& To SendMessage(rList&, LB_GETCOUNT, 0&, 0&) - 1&
             ScreenName$ = String$(4&, vbNullChar)
             itmHold& = SendMessage(rList&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
             itmHold& = itmHold& + 24&
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4&, rBytes&)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4&)
             psnHold& = psnHold& + 6&
             ScreenName$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ' screenname$ = Left$(screenname$, InStr(screenname$, vbNullString) - 1&)
            thecontrol.AddItem RemSpaces(ScreenName$)
        Next index&
        Call CloseHandle(mThread&)
    End If
End Sub
Public Function aimsnfromim() As String
    Dim lngimwin As Long, strimcap As String, lngtextlen As Long, strsn As String
     lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& = 0& Then
        Exit Function
    Else
         strimcap$ = GetCaption(lngimwin&)
        If InStr(strimcap$, "- instant message") <> 0& Then
             lngtextlen& = GetWindowTextLength(lngimwin&) - 19&
             strsn$ = Left$(strimcap$, InStr(strimcap$, "" & strimcap$ & "") + lngtextlen&)
             aimsnfromim$ = strsn$
        Else
             aimsnfromim$ = ""
        End If
    End If
End Function

Public Function aimusersn() As String
    Dim lngbuddywin As Long, thecap As String
     lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
     thecap$ = GetCaption(lngbuddywin&)
    If Not Right(thecap$, 20&) = "'s Buddy List Window" Then
         aimusersn$ = ""
        Exit Function
    Else
         aimusersn$ = Left$(thecap$, Len(thecap$) - 20&)
    End If
End Function


Public Sub antiidle()
    Dim lngpate As Long, lngmodal As Long
    Dim lngbutton As Long, lngbutton2 As Long
     lngpate& = FindWindow("_aol_pate", vbNullString)
     lngmodal& = FindWindow("_aol_modal", vbNullString)
     lngbutton& = FindWindowEx(lngpate&, 0&, "_aol_icon", vbNullString)
     lngbutton2& = FindWindowEx(lngmodal&, 0&, "_aol_icon", vbNullString)
    Call SendMessageLong(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngbutton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessageLong(lngbutton2&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngbutton2&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aolautosncollect(TheList As Control, lnghowmanysns As Long)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngchild As Long, lnglist1 As Long, lnglist As Long
    Dim lngicon1 As Long, lngicon2 As Long, lngicon3 As Long, lngicon4 As Long, lngicon5 As Long
    Dim lngicon6 As Long, lngicon7 As Long, lngicon8 As Long, lngchatting As Long, lngcounter As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    Pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Find a Chat")
        Let lnglist1& = FindWindowEx(lngchild&, 0&, "_AOL_Listbox", vbNullString)
        Let lnglist& = FindWindowEx(lngchild&, lnglist&, "_AOL_Listbox", vbNullString)
        Let lngicon1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon2& = FindWindowEx(lngchild&, lngicon1&, "_AOL_Icon", vbNullString)
        Let lngicon3& = FindWindowEx(lngchild&, lngicon2&, "_AOL_Icon", vbNullString)
        Let lngicon4& = FindWindowEx(lngchild&, lngicon3&, "_AOL_Icon", vbNullString)
        Let lngicon5& = FindWindowEx(lngchild&, lngicon4&, "_AOL_Icon", vbNullString)
        Let lngicon6& = FindWindowEx(lngchild&, lngicon5&, "_AOL_Icon", vbNullString)
        Let lngicon7& = FindWindowEx(lngchild&, lngicon6&, "_AOL_Icon", vbNullString)
        Let lngicon8& = FindWindowEx(lngchild&, lngicon7&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngchild&, lngicon8&, "_AOL_Icon", vbNullString)
        Pause 1
    Loop Until lngchild& <> 0& And lnglist1& <> 0& And lnglist& <> 0& And lnglist1& <> lnglist& And lngicon& <> 0& And lngicon1& <> 0& And lngicon2& <> 0& And lngicon3& <> 0& And lngicon4& <> 0& And lngicon5& <> 0& And lngicon6& <> 0& And lngicon7& <> 0& And lngicon8& <> 0& And _
    lngicon1& <> lngicon2& And lngicon2& <> lngicon3& And lngicon3& <> lngicon4& And lngicon4& <> lngicon5& And lngicon5& <> lngicon6& And lngicon6& <> lngicon7& And lngicon7& <> lngicon8& And lngicon8& <> lngicon&
    MsgBox "Yuo have 10 seconds to pick a category"
    timeout 9
    Call waitforlisttoload(lnglist&)
    Do: DoEvents
        Call clicklistindex(lnglist&, lngcounter&)
        Call ClickIcon(lngicon&)
        Let lngcounter& = lngcounter& + 1&
        Do: DoEvents
            Let lngchatting& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Who's Chatting")
            Let lnglist1& = FindWindowEx(lngchatting&, 0&, "_AOL_Listbox", vbNullString)
        Loop Until lngchatting& <> 0& And lnglist1& <> 0&
        Call waitforlisttoload(lnglist1&)
        Call addlisttocontrol(lnglist1&, TheList)
        Call PostMessage(lngchatting&, WM_CLOSE, 0&, 0&)
    Loop Until TheList.ListCount > lnghowmanysns&
    Call SetCursorPos(cursorpos.x, cursorpos.Y)
End Sub


Public Sub aimchatignorebyname(strwho$)
    Dim lngchatwin As Long, lngicon&, TempString As String, tempstring2 As String
    Dim index As Long, strBuffer As String, lngtree As Long
    Dim lngitemdata As Long, lngtextlen As Long, strText As String
     lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
     lngtree& = FindWindowEx(lngchatwin&, 0&, "_oscar_tree", vbNullString)
    If lngtree& = 0& Then Exit Sub
    For index& = 1& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
         lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
         strBuffer$ = String(lngtextlen&, 0&)
         strText$ = SendMessageByString(lngtree&, LB_GETtext, index&, ByVal strBuffer$)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
         lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then
             TempString$ = ReplaceString(strBuffer$, " ", "")
             tempstring2$ = ReplaceString(strwho$, " ", "")
            If LCase$(TempString$) = LCase$(tempstring2$) Then
                 lngicon& = FindWindowEx(lngchatwin&, 0&, "_oscar_iconbtn", vbNullString)
                 lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
                Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
            End If
        End If
    Next index&
End Sub
Public Sub aimcloseim(strwho As String)
    Dim lngimwin As Long, TempString As String, tempstring2 As String
     lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& <> 0& Then
         TempString$ = GetCaption(lngimwin&)
         TempString$ = LCase(ReplaceString(TempString$, " ", ""))
         tempstring2$ = LCase(ReplaceString(strwho$, " ", "") & "-instantmessage")
        If TempString$ Like tempstring2$ Then
            Call windowclose(lngimwin&)
        Else
            Do: DoEvents
                 lngimwin& = FindWindowEx(0&, lngimwin&, "aim_imessage", vbNullString)
                 TempString$ = GetCaption(lngimwin&)
                 TempString$ = ReplaceString(TempString$, " ", "")
                 tempstring2$ = ReplaceString(strwho$, " ", "") & "-instantmessage"
                If TempString$ Like tempstring2$ Then
                    Call windowclose(lngimwin&)
                End If
            Loop Until lngimwin& = 0 Or TempString$ Like tempstring2$
        End If
    End If
End Sub

Public Sub aimenterroom(strroom As String)
    Dim lngchatwin As Long, strstring As String
    Call aimsendchatinvite(aimusersn$, aimusersn$, strroom$)
    Do: DoEvents
         lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Loop Until lngchatwin& <> 0
    Do: DoEvents
         strstring$ = GetCaption(lngchatwin&)
    Loop Until Right$(strstring$, 5&) = Right$(strroom$, 5&)
End Sub

Public Function aimfindroom(strName As String) As Long
    Dim lngchatwin As Long
     lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    If lngchatwin& <> 0& Then
        If LCase$(GetCaption(lngchatwin&)) Like LCase$(strName$) Then
             aimfindroom = lngchatwin&
        Else
            Do: DoEvents
                 lngchatwin& = FindWindowEx(0&, lngchatwin&, "aim_chatwnd", vbNullString)
                If LCase$(GetCaption(lngchatwin&)) Like LCase$(strName$) Then
                     aimfindroom = lngchatwin&
                End If
            Loop Until LCase$(GetCaption(lngchatwin&)) Like LCase$(strName$) Or lngchatwin& = 0&
        End If
    End If
End Function
Public Function aimisbuddyavailable() As Boolean
    Dim lngbuddywin As Long, lnglocatewin As Long
     lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "get member inf&o")
    Do: DoEvents
         lnglocatewin& = FindWindow("_oscar_locate", vbNullString)
    Loop Until lnglocatewin& <> 0&
End Function

Public Function aimisonline() As Boolean
    Dim lngbuddywin As Long
     lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    If lngbuddywin& <> 0& Then aimisonline = True: Exit Function
     aimisonline = False
End Function

Public Sub aimkeyword(TheKeyWord$)
    Dim lngbuddywin As Long, lngeditwin As Long, lngicon As Long
     lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
     lngeditwin& = FindWindowEx(lngbuddywin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, TheKeyWord$)
     lngicon& = FindWindowEx(lngbuddywin&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aimopenim()
    Dim lngbuddy As Long, lnggroup As Long, lngicon As Long, lngimwin As Long
     lngbuddy& = FindWindow("_oscar_buddylistwin", vbNullString)
     lnggroup& = FindWindowEx(lngbuddy&, 0&, "_oscar_tabgroup", vbNullString)
     lngicon& = FindWindowEx(lnggroup&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngimwin& = FindWindow("aim_imessage", vbNullString)
    Loop Until lngimwin& <> 0&
End Sub


 

Public Sub aimsendchat(strmessage As String)
    Dim lngchatwin As Long, lngatewin As Long, lngicon As Long
     lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
     lngatewin& = FindWindowEx(lngchatwin&, 0&, "wndate32class", vbNullString)
     lngatewin& = FindWindowEx(lngchatwin&, lngatewin&, "wndate32class", vbNullString)
    Call SendMessageByString(lngatewin&, WM_SETTEXT, 0, strmessage$)
     lngicon& = FindWindowEx(lngchatwin&, 0&, "_oscar_iconbtn", vbNullString)
     lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
     lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
     lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aimsendchatinvite(strwho As String, strmessage As String, strchatname As String)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngicon As Long, lnginvitewin As Long
    Dim lngeditwin As Long, lngchatwin As Long
     lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
     lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
     lngicon& = FindWindowEx(lnggroupwin&, 0&, "_oscar_iconbtn", vbNullString)
     lngicon& = FindWindowEx(lnggroupwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do While lnginvitewin = 0& And lngeditwin& = 0& And lngicon& = 0&
         lnginvitewin& = FindWindow("aim_chatinvitesendwnd", vbNullString)
         lngeditwin& = FindWindowEx(lnginvitewin&, 0&, "edit", vbNullString)
         lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
         lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
         lngicon& = FindWindowEx(lnginvitewin&, 0&, "_oscar_iconbtn", vbNullString)
         lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
         lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
        DoEvents
    Loop
     lngeditwin& = FindWindowEx(lnginvitewin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strwho$)
    DoEvents
     lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strmessage$)
    DoEvents
     lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strchatname$)
    DoEvents
     lngicon& = FindWindowEx(lnginvitewin&, 0&, "_oscar_iconbtn", vbNullString)
     lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
     lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Loop Until lngchatwin& <> 0&
    If FindWindow("#32770", "Buddy Chat Invitation Error") Then Call windowclose(FindWindow("#32770", "Buddy Chat Invitation Error"))
End Sub

Public Sub aimsendchatlink(strurl As String, strmessage As String)
    Call aimsendchat("< a href=" & Chr$(34&) & strurl$ & Chr$(34&) & ">" & strmessage$ & "</a>")
End Sub

Public Sub aimsendchattoallrooms(Message As String)
    Dim lngchatwin As Long
     lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    If lngchatwin& <> 0& Then
        Do: DoEvents
            lngchatwin& = FindWindowEx(0&, lngchatwin&, "aim_chatwnd", vbNullString)
            If lngchatwin& <> 0& Then Call aimsendchat(Message$)
        Loop Until lngchatwin& = 0&
    End If
End Sub

Public Sub aimsendim(strsn As String, strmessage As String)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngbutton As Long, lngimwin As Long
    Dim lngcombowin As Long, lngeditwin As Long, lngatewin&, lngicon As Long
    Dim lngerrorwin As Long, TempString As String, tempstring2 As String
     lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& <> 0 Then
         TempString$ = GetCaption(lngimwin&)
         TempString$ = LCase$(ReplaceString(TempString$, " ", ""))
         tempstring2$ = LCase$(ReplaceString(strsn$, " ", "") & "-instantmessage")
        If TempString$ Like tempstring2$ Then
             lngatewin& = FindWindowEx(lngimwin&, 0&, "wndate32class", vbNullString)
             lngatewin& = FindWindowEx(lngimwin&, lngatewin&, "wndate32class", vbNullString)
            Call SendMessageByString(lngatewin&, WM_SETTEXT, 0, strmessage$)
            DoEvents
             lngicon& = FindWindowEx(lngimwin&, 0&, "_oscar_iconbtn", vbNullString)
            Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
            Exit Sub
        End If
    End If
     lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
     lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
     lngbutton& = FindWindowEx(lnggroupwin&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngimwin& = FindWindow("aim_imessage", vbNullString)
    Loop Until lngimwin& <> 0&
     lngcombowin& = FindWindowEx(lngimwin&, 0&, "_oscar_persistantcombo", vbNullString)
     lngeditwin& = FindWindowEx(lngcombowin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strsn$)
    DoEvents
     lngatewin& = FindWindowEx(lngimwin&, 0&, "wndate32class", vbNullString)
     lngatewin& = FindWindowEx(lngimwin&, lngatewin&, "wndate32class", vbNullString)
    Call SendMessageByString(lngatewin&, WM_SETTEXT, 0, strmessage$)
    DoEvents
     lngicon& = FindWindowEx(lngimwin&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub aimsendmassim(thepeople As Listbox, strmessage As String)
    Dim index As Long
    If thepeople.ListCount = 0& Then Exit Sub
    For index& = 0& To thepeople.ListCount - 1&
        Call aimsendim(thepeople.List(index&), strmessage$)
        Call Pause(3)
    Next index&
End Sub
Public Sub aimsignoff()
    Call runanymenubystring(FindWindow("_oscar_buddylistwin", vbNullString), "sign o&ff")
End Sub

Public Sub aimsignon(stryoursn As String, stryourpassword As String)
    Dim isopened As Boolean, lngcombo As Long, lngaimwin As Long, lngedit As Long, lngstatic As Long
    On Error Resume Next
    If aimisonline Then Exit Sub
    If file_ifileexists("c:\program files\aim95\aim.exe") = True Then
        Call Shell("c:\program files\aim95\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\program files\netscape\communicator\program\aim\aim.exe") = True Then
        Call Shell("c:\program files\netscape\communicator\program\aim\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\aim\aim.exe") Then
        Call Shell("c:\aim\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\aim95\aim.exe") Then
        Call Shell("c:\aim95\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\aim2\aim.exe") Then
        Call Shell("c:\aim2\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\program files\aim2\aim.exe") Then
        Call Shell("c:\program files\aim2\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\online\aim\aim.exe") Then
        Call Shell("c:\online\aim\aim.exe", vbNormalFocus)
            ElseIf file_ifileexists("c:\program files\aim98\aim.exe") Then
        Call Shell("c:\program files\aim98\aim.exe", vbNormalFocus)
    Else
        Exit Sub
    End If
    'i just realized, selectcase would have been a little bit easier there, damn.
    Do: DoEvents
         lngaimwin& = FindWindow("#32770", "sign on")
    Loop Until lngaimwin& <> 0&
     lngcombo& = FindWindowEx(lngaimwin&, 0&, "combobox", vbNullString)
     lngedit& = FindWindowEx(lngcombo&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, stryoursn$): DoEvents
     lngedit& = FindWindowEx(lngaimwin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, stryourpassword$)
    Call PostMessage(lngedit&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngedit&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub addorremovebuddy(strName As String, strgroup, booladdbuddy As Boolean)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
     strrealgroup$ = strgroup
     strgroup = ReplaceString(strgroup, " ", "")
     strName$ = ReplaceString(strName$, " ", "")
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call Keyword("buddy view")
    End If
    Do: DoEvents
         lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
         lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
         lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
         lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s buddy lists")
         lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
         lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
         lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until lngsetupwin& <> 0& And lnggroups& <> 0& And lngeditbut& <> 0&
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&)
         strstring$ = getlistitemtext(lnggroups&, lngindex&)
         strstring$ = ReplaceString(Left$(LCase$(strstring$), InStr(strstring, Chr$(9&)) - 1&), " ", "")
        If strstring$ = strgroup Then
            Exit For
        End If
    Next lngindex&
    If lngindex& = -1& Then
        Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
        Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    Else
         lngtab& = InStr(strstring$, Chr$(9&))
        If lngtab& <> 0& Then strstring$ = Left$(strstring$, lngtab& - 1&)
        Call SendMessageLong(lnggroups&, LB_SETCURSEL, CLng(lngindex&), 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lnggroupwin& = FindWindowEx(lngmdi&, 0&, "aol child", "edit list " & strrealgroup$)
             lngaddbox& = FindWindowEx(lnggroupwin&, 0&, "_aol_edit", vbNullString)
             lngaddbox& = FindWindowEx(lnggroupwin&, lngaddbox&, "_aol_edit", vbNullString)
             lngblist& = FindWindowEx(lnggroupwin&, 0&, "_aol_listbox", vbNullString)
             lngaddbut& = FindWindowEx(lnggroupwin&, 0&, "_aol_icon", vbNullString)
             lngremovebut& = FindWindowEx(lnggroupwin&, lngaddbut&, "_aol_icon", vbNullString)
             lngsavebut& = FindWindowEx(lnggroupwin&, lngremovebut&, "_aol_icon", vbNullString)
             lngstatic& = FindWindowEx(lnggroupwin&, 0&, "_aol_static", vbNullString)
             lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
             lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
             lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
             lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
        Loop Until lnggroupwin& <> 0& And lngaddbox& <> 0& And lngblist <> 0& And lngaddbut <> 0& And lngremovebut& <> 0& And lngsavebut& <> 0& And lngstatic& <> 0&
        Call waitforlisttoload(lngblist&)
        If booladdbuddy = True Then
            If getlistitemindex(lngblist&, LCase$(ReplaceString(strName$, " ", ""))) <> -1& Then
                Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
                Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
                Exit Sub
            End If
            Call SendMessageByString(lngaddbox&, WM_SETTEXT, 0&, strName$)
            Call SendMessageLong(lngaddbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngaddbut&, WM_LBUTTONUP, 0&, 0&)
            Do: DoEvents
                 lngindex& = getlistitemindex(lngblist&, LCase$(ReplaceString(strName$, " ", "")))
                 strstring$ = GetText(lngstatic&) ' meaning buddy is already there !
            Loop Until lngindex& <> 0& Or strstring$ <> " "
            If strstring$ <> " " Then
                Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
            ElseIf strstring$ = " " Then
                Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
            End If
        ElseIf booladdbuddy = False Then
             lngbindex& = getlistitemindex(lngblist&, LCase$(ReplaceString(strName$, " ", "")))
            If lngbindex& = -1& Then
                Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
                Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
                Exit Sub
            End If
            Call SendMessageLong(lngblist&, LB_SETCURSEL, CLng(lngbindex&), 0&)
            Call SendMessageLong(lngremovebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngremovebut&, WM_LBUTTONUP, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
        End If
        Do Until lngokwin& <> 0& And lngokbut& <> 0&
             lngokwin& = FindWindow("#32770", "america online")
             lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "ok")
        Loop: DoEvents
        Do Until lngokwin& = 0&
            Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
             lngokwin& = FindWindow("#32770", "america online")
        Loop: DoEvents
        Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub y2kcheck()
  Dim vbNewLine
    Dim dwLCID As Long
    dwLCID = GetSystemDefaultLCID()
    If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, "MM/dd/yyyy") = False Then
       MsgBox "Failed"
       Exit Sub
    End If
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
    MsgBox "The Windows Date Format is now fixed from Y2K Threat, this does not include ""Software and addon programs, problems may occour with them.""(", vbInformation, "Y2K Compliant"
End Sub

Public Function CryptString2(txtString As String, Encrypt As Boolean) As String
       On Error GoTo errhandler
       Dim x As Integer
       Dim outString As String
       Dim iLen As Integer
       Dim sFirstSeed As String
       Dim sSecondSeed As String
       Dim iSeed As Integer
       If Encrypt Then
           sFirstSeed = Left(txtString, 1)
           sSecondSeed = Mid(txtString, 2, 1)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString)
           For x = 1 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) + 2) & outString
           Next
           outString = Chr(Asc(sFirstSeed) * 2 + 3) & outString
           outString = outString & Chr(Asc(sSecondSeed) * 2 - 3)
       Else
           sFirstSeed = Chr((Asc(Left(txtString, 1)) - 3) \ 2)
           sSecondSeed = Chr((Asc(Right(txtString, 1)) + 3) \ 2)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString) - 1
           For x = 2 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) - 2) & outString
           Next
       End If
       CryptString2 = outString
       Exit Function
errhandler:
       MsgBox "Error in CryptString" & vbCrLf & "Error: " & Err.Description & vbCrLf & "Number: " & Err.number
       CryptString2 = ""
   End Function


Sub sys_FontsToList(lst As Listbox)
Dim i
For i = 0 To Screen.FontCount - 1  ' Determine number of fonts.
    lst.AddItem Screen.Fonts(i)  ' Put each font into list box.
Next i
End Sub

Public Sub picmenu(subm, mainm, frmy As Form, picto As PictureBox)
Dim hMenu&
Dim hSubMenu&
Dim hID&
hMenu& = GetMenu(frmy.hwnd)
hSubMenu& = GetSubMenu(hMenu&, subm)
hID& = GetMenuItemID(hSubMenu&, mainm)
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, picto.Picture, picto.Picture

End Sub




Sub timeout(Duration)
Dim starttime
Dim x
starttime = Timer
Do While Timer - starttime < Duration
x = DoEvents()
Loop
End Sub







Sub sys_shutdown(ByVal Xit As Integer)
Dim n
Select Case Xit
    Case 1
        n = ExitWindowsEx(67, 0) 'reboot the computer
    Case 2
        n = ExitWindowsEx(66, 0) 'restart Windows
    Case 3
        n = ExitWindowsEx(0, 0) 'exit Windows
End Select
End Sub


Sub text_charcount(Text1 As TextBox)
Dim asdf
asdf = Len(Text1)
MsgBox "there are " + asdf + " characters"

End Sub

Sub text_findword(Text1 As TextBox, Txt As String)
Dim x%
x% = InStr(Text1.Text, Txt$)
If x% <> 0 Then
Text1.SelStart = x% - 1
Text1.SelLength = Len(Txt$) ' or 5
Text1.SetFocus
End If

End Sub


Function trimnull(ByVal i) As String
Dim a
Dim t$
For a = 1 To Len(i)
    If (Mid$(i, a, 1) <> Chr$(0)) Then
     t$ = t$ + Mid$(i, a, 1)
    Else
     trimnull = t$
     Exit Function
    End If
Next
trimnull = t$
End Function

Function trm(ByVal t As String)
Dim a As Integer
Dim b As String
Dim c As String
For a = 1 To Len(t)
b = Mid(t, a, 1)
If b = Chr(0) Then b = ""
c = c & b
Next
trm = c
End Function



Sub web_open2()
'Sub declared by Tech
Dim a!
a! = Shell("explorer http://www.absolutezer0.com")
End Sub


Sub web_openpage()
Dim openit!
Dim NoFreeze
openit! = Shell("start HTTP://www.absolutezer0.com"): NoFreeze = DoEvents()
End Sub


Public Sub win98help()
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Dim lngrunform As Long, lngcombo As Long
     lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
     lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
     lngbasebar& = FindWindow("BaseBar", vbNullString)
     lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
     lngtoolwin& = FindWindowEx(lngmenusite&, 0&, "ToolbarWindow32", vbNullString)
    Pause 0.2
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyH, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyH, 0&)
End Sub

Public Sub win98run(strFile As String)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Dim lngrunform As Long, lngcombo As Long
     lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
     lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
     lngbasebar& = FindWindow("BaseBar", vbNullString)
     lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
     lngtoolwin& = FindWindowEx(lngmenusite&, 0&, "ToolbarWindow32", vbNullString)
    Pause 0.2
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyR, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyR, 0&)
    Do: DoEvents
         lngrunform& = FindWindow("#32770", vbNullString)
         lngcombo& = FindWindowEx(lngrunform&, 0&, "ComboBox", vbNullString)
         lngcombo& = FindWindowEx(lngcombo&, 0&, "Edit", vbNullString)
         lngokbut& = FindWindowEx(lngrunform&, 0&, "Button", vbNullString)
    Loop Until lngrunform& <> 0& And lngcombo& <> 0&
    Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub win98runstartmenu(lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
     lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
     lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
     lngbasebar& = FindWindow("BaseBar", vbNullString)
     lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
     lngmenusite& = GetMenu(lngmenusite&)
    Call runanymenu(lngmenusite&, lngmenunumber&, lngsubmenunumber&)
End Sub

Public Sub win98runstartmenubystring(strmenutext As String)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
     lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
     lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
     lngbasebar& = FindWindow("BaseBar", vbNullString)
     lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
     lngmenusite& = GetMenu(lngmenusite&)
    Call runanymenubystring(lngmenusite&, strmenutext$)
End Sub

Function AnyPadsLeft() As Integer
    Dim i As Integer
Dim Document
Dim FState
    ' Cycle through the document array.
    ' Return true if there is at least one open document.
    For i = 1 To UBound(Document)
        If Not FState(i).Deed Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next
End Function



Public Sub win98shutdown()
    Static ewx_shutdown
    Dim getmsg As Long
     getmsg& = MsgBox("are you sure you want to exit windows?", vbYesNo Or vbQuestion)
    If getmsg& = vbNo Then
        Exit Sub
    Else
        Call ExitWindowsEx(ewx_shutdown, 0&)
    End If
End Sub

Public Sub winampaddsongtoplaylist(strsong As String)
    Dim lngopenwin As Long, lngedit As Long, lngopenbut As Long
    Call PostMessage(winampplaylistwin&, WM_KEYDOWN, vbKeyL, 0&)
    Call PostMessage(winampplaylistwin&, WM_KEYUP, vbKeyL, 0&)
    Do: DoEvents
         lngopenwin& = FindWindow("#32770", "Open file(s)")
         lngedit& = FindWindowEx(lngopenwin&, 0&, "Edit", vbNullString)
         lngopenbut& = FindWindowEx(lngopenwin&, 0&, "Button", "&Open")
        Pause 1
    Loop Until lngopenwin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strsong$)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function winampequalizerwin() As Long
     winampequalizerwin& = FindWindow("Winamp EQ", vbNullString)
End Function

Public Sub winampjumptosongbyindex(index As Long)
    Dim lngjumpwin As Long, lnglist As Long, strText As String
     lngjumpwin& = FindWindow("#32770", "Jump to file")
     lnglist& = FindWindowEx(lngjumpwin&, 0&, "ListBox", vbNullString)
    Call SendMessageLong(lnglist&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(lnglist&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lnglist&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub winampjumptosongbyname(strName As String)
    Dim lngjumpwin As Long, lnglist As Long, strText As String
     lngjumpwin& = FindWindow("#32770", "Jump to file")
     lnglist& = FindWindowEx(lngjumpwin&, 0&, "ListBox", vbNullString)
    Call SendMessageLong(lnglist&, LB_SETCURSEL, gettreeitemindex(lnglist&, strName$), 0&)
    Call PostMessage(lnglist&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lnglist&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub winampnextsong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyB, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyB, 0&)
End Sub

Public Sub winampopenskins()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyControl And vbKeyL, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyControl And vbKeyL, 0&)
End Sub

Public Sub winampopensong(strsong As String)
    Dim lngopenwin As Long, lngedit As Long, lngopenbut As Long
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyL, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyL, 0&)
    Do: DoEvents
         lngopenwin& = FindWindow("#32770", "Open file(s)")
         lngedit& = FindWindowEx(lngopenwin&, 0&, "Edit", vbNullString)
         lngopenbut& = FindWindowEx(lngopenwin&, 0&, "Button", "&Open")
        Pause 1
    Loop Until lngopenwin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strsong$)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub winamppausesong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyC, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyC, 0&)
End Sub

Public Function winampplaylistwin() As Long
     winampplaylistwin& = FindWindow("Winamp PE", vbNullString)
End Function

Public Sub winampprevioussong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyZ, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyZ, 0&)
End Sub

Public Sub winampstartplay()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyX, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyX, 0&)
End Sub

Public Sub winampstopsong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyV, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyV, 0&)
End Sub

Public Sub winampturndownvolume(lngdecibals As Long)
    Dim index As Long
    For index& = 0& To lngdecibals&
        Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyDown, 0&)
        Call PostMessage(winampwin&, WM_KEYUP, vbKeyDown, 0&)
    Next index&
End Sub

Public Sub winampturnupvolume(lngdecibals As Long)
    Dim index As Long
    For index& = 0& To lngdecibals&
        Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyUp, 0&)
        Call PostMessage(winampwin&, WM_KEYUP, vbKeyUp, 0&)
    Next index&
End Sub

Public Function winampwin() As Long
     winampwin& = FindWindow("Winamp v1.x", vbNullString)
End Function

Public Sub sendmassim(strwho As Control, strmessage As String)
    Dim lngindex As Long
    For lngindex& = 0& To strwho.ListCount - 1&
        Call InstantMessage(strwho.List(lngindex&), strmessage$)
        Pause 0.7
    Next lngindex&
End Sub
Public Sub windowclose(lnghwnd As Long)
    Call SendMessageLong(lnghwnd&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub WindowHide(lnghwnd As Long)
    Call ShowWindow(lnghwnd&, SW_HIDE)
End Sub

Public Sub windowmaximize(lnghwnd As Long)
    Call ShowWindow(lnghwnd, SW_MAXIMIZE)
End Sub

Public Sub WindowShow(lnghwnd As Long)
    Call ShowWindow(lnghwnd&, SW_SHOW)
End Sub

Public Sub writeini(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub

Function Definition(Word As String)
Dim Source As String, Link As String, Start As Integer, Lstart As Integer, Temp As String

Source = GetSource("http://www.allwords.com/query.php?SearchType=3&Keyword=" & Word & "&goquery=Find+it%21&Language=ENG")
Link = Source

Start = (InStr(1, Source, "</a></b></font></td><td v_align='top'><font color='BLACK' face='Tahoma'>", vbTextCompare) + 72)
Lstart = (InStr(1, Link, "<img src='spacer.gif' width='1' height='1'></td></tr><tr valign='top'><td>", vbTextCompare) + 74)

Source = Mid(Source, Start, Len(Source) - Start)
Link = Mid(Link, Lstart, Len(Link) - Lstart)

Temp = Mid(Link, 1, InStr(1, Link, ">", vbTextCompare))
Temp = Replace(Temp, "href='", "href=http://www.allwords.com/", 1, Len(Temp), vbTextCompare)
Temp = Replace(Temp, "'", "", 1, Len(Temp), vbTextCompare)
Definition = color1 & "[" & Temp & color2 & Word & "</a>" & color1 & "] - " & Mid(Source, 1, InStr(1, Source, ".", vbTextCompare))
Definition = Replace(Definition, "<HR></td></tr><tr valign='top'><td>", "", 1, Len(Definition), vbTextCompare)
If InStr(1, Definition, "<title>", vbTextCompare) <> 0 Then Definition = Definition2(Word) '"Definition not found"
End Function

Function Definition2(Word As String)
On Error GoTo errror
Dim Source As String, Start As Integer
Source = GetSource("http://lookwayup.com/lwu.exe//lwu/d?lang=Non&w=" & Word)
Start = (InStr(1, Source, "1.</FONT></B></TD><TD CLASS=""TEXT""> [", vbTextCompare) + 42)
Source = Mid(Source, Start, Len(Source) - Start)
Definition2 = color1 & "[" & color2 & Word & color1 & "] - " & Mid(Source, 1, InStr(1, Source, ". <A HREF=", vbTextCompare) - 1)
'Definition2 = Replace(Definition2, "<br>", "", 1, Len(Definition2), vbTextCompare)
'If InStr(1, Synonyms, "<title>", vbTextCompare) <> 0 Then Definition2 = "Definition not found"
Exit Function
errror:
Definition2 = color1 & "[" & color2 & Word & color1 & "] Definition not found"
End Function

Function GetDialerStats()
'On Error GoTo errror
Dim Source As String, Start As Integer, User As String, PW As String
Dim i As Long, HTMLElement
User = Left(DialerLogin, InStr(1, DialerLogin, ":", vbTextCompare) - 1)
PW = Right(DialerLogin, Len(DialerLogin) - InStr(1, DialerLogin, ":", vbTextCompare))

Form1.WebBrowser1.Navigate ("http://www.statsbank.com/login.php?frmuser=" & User & "&frmpass=" & PW)
Do
DoEvents
Loop Until Form1.WebBrowser1.ReadyState = 4
Form1.WebBrowser1.Navigate ("http://www.statsbank.com/summary_total.php")
Do
DoEvents
Loop Until Form1.WebBrowser1.ReadyState = 4

'For i = 1 To Form1.WebBrowser1.Document.All.Length
        Set HTMLElement = Form1.WebBrowser1.Document.All.Item(1)
        Form1.Text3 = HTMLElement.innerHTML
        'Next i
        
'Form1.Text3 = "Sdfas"
'Form1.Text3 = Source
'Definition2 = color1 & "[" & color2 & Word & color1 & "] Definition not found"
End Function

Function WordOfTheDay()
On Error GoTo errror
Dim Source As String, Start As Integer, Word As String
Source = GetSource("http://dictionary.reference.com/wordoftheday/")
Start = (InStr(1, Source, "Day: ", vbTextCompare) + 5)
Source = Mid(Source, Start, Len(Source) - Start)
Word = Mid(Source, 1, InStr(1, Source, "</title>", vbTextCompare) - 1)
WordOfTheDay = Definition(Word)
'Definition2 = Replace(Definition2, "<br>", "", 1, Len(Definition2), vbTextCompare)
'If InStr(1, Synonyms, "<title>", vbTextCompare) <> 0 Then Definition2 = "Definition not found"
Exit Function
errror:
WordOfTheDay = color1 & "[" & color2 & Word & color1 & "] Definition not found"
End Function
Function Synonyms(Word As String)
Dim Source As String, Start As Integer
Source = GetSource("http://thesaurus.reference.com/search?q=" & Word)
Start = (InStr(1, Source, "<tr><td align=right valign=top><b>Synonyms:</b>&nbsp;&nbsp;</td><td>", vbTextCompare) + 68)
Source = Mid(Source, Start, Len(Source) - Start)
Synonyms = "[Synonyms for " & Word & "] - " & Mid(Source, 1, InStr(1, Source, "</td></tr>", vbTextCompare) - 1)
Synonyms = Replace(Synonyms, "<b style=""background: #ffffaa"">", "", 1, Len(Synonyms), vbTextCompare)
If InStr(1, Synonyms, "<title>", vbTextCompare) <> 0 Then Synonyms = "Synonyms not found"
End Function

Function Antonyms(Word As String)
Dim Source As String, Start As Integer
Source = GetSource("http://thesaurus.reference.com/search?q=" & Word)
Start = (InStr(1, Source, "<tr><td align=right valign=top><b>Antonyms:</b>&nbsp;&nbsp;</td><td>", vbTextCompare) + 68)
Source = Mid(Source, Start, Len(Source) - Start)
Antonyms = "[Antonyms for " & Word & "] - " & Mid(Source, 1, InStr(1, Source, "</td></tr>", vbTextCompare) - 1)
Antonyms = Replace(Antonyms, "<b style=""background: #ffffaa"">", "", 1, Len(Antonyms), vbTextCompare)
If InStr(1, Antonyms, "<title>", vbTextCompare) <> 0 Then Antonyms = "Antonyms not found"
End Function
Function Horoscope(sign As String)
Dim Source As String, Start As Integer
Select Case LCase(sign)
Case "pisces"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/piscesdailyhoroscope.html")
Case "cancer"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/cancerdailyhoroscope.html")
Case "aquarius"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/aquariusdailyhoroscope.html")
Case "libra"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/libradailyhoroscope.html")
Case "gemini"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/geminidailyhoroscope.html")
Case "taurus"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/taurusdailyhoroscope.html")
Case "virgo"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/virgodailyhoroscope.html")
Case "capricorn"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/capricorndailyhoroscope.html")
Case "aries"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/ariesdailyhoroscope.html")
Case "leo"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/leodailyhoroscope.html")
Case "sagittarius"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/sagittariusdailyhoroscope.html")
Case "scorpio"
Source = GetSource("http://astrology.yahoo.com/us/astrology/today/scorpiodailyhoroscope.html")
Case Else
Horoscope = "Invalid Sign": GoTo Alldone
End Select
Start = (InStr(1, Source, "<font face=""Arial"" size=""-1"" color=black>", vbTextCompare) + 41)
Source = Mid(Source, Start, Len(Source) - Start)
Horoscope = Mid(Source, 1, InStr(1, Source, "</font></TD></TR></table>", vbTextCompare) - 1)
Alldone:
End Function
Function SetCDPlayerReady()
 mciSendString "open cdaudio Alias cd wait shareable", 0, 0, 0
End Function


Sub scanner(J$)
Dim nm%
' cih simple scan
Dim h%, w%
nm% = nm% + 1
Static O$
h% = 0
Dim Y!, Z!
Dim Q%, F%
w% = 32000
F = FreeFile
Open J$ For Binary As F
Z! = LOF(F)
If Z! < w% Then w% = Z!
Y! = Z! - w%
'Do: n% = DoEvents()
Z! = LOF(F)
If Y! < w% Then
Q% = Y!
Else
Q% = w%
End If
O$ = String(Q, " ")
Get F, , O$
If InStr(LCase$(O$), "cih v1.2 ttit") Then h% = 1
If InStr(LCase$(O$), "cih v1.3 ttit") Then h% = 2
If InStr(LCase$(O$), "cih v1.4 tatung") Then
Y! = Y! - Q%
'Loop Until y! <= 0 Or h% = 1
Close F
If h% = 1 Then
 'sp.List1.AddItem "cih v1.2" & Chr(9) & sp.File1.Path & "\" & sp.File1.List(br%)
ElseIf h% = 2 Then
 'sp.List1.AddItem "cih v1.3" & Chr(9) & sp.File1.Path & "\" & sp.File1.List(br%)
ElseIf h% = 3 Then
 'sp.List1.AddItem "cih v1.4" & Chr(9) & sp.File1.Path & "\" & sp.File1.List(br%)
End If
End If
End Sub

Sub sendkeystrokes()
'server

'send this data

'winsock1.senddata VbCrlf
'winsock1.senddata text1.text

'end sub

'--------------------------

'client

'on data arriaval

'Dim A As String
'winsock1.getdata a,vbstring
'end sub

'SendKeys (A)

End Sub

Function SoundDecode(ByVal InTXT As String) As String
Dim t
Dim T_SoundDecode$
InTXT$ = LCase(InTXT$ & " ")
InTXT$ = AF_Script(InTXT$, "aigh ", "a")
InTXT$ = AF_Script(InTXT$, "bee ", "b")
InTXT$ = AF_Script(InTXT$, "see ", "c")
InTXT$ = AF_Script(InTXT$, "dee ", "d")
InTXT$ = AF_Script(InTXT$, "eff ", "f")
InTXT$ = AF_Script(InTXT$, "gee ", "g")
InTXT$ = AF_Script(InTXT$, "aytch ", "h")
InTXT$ = AF_Script(InTXT$, "eye ", "i")
InTXT$ = AF_Script(InTXT$, "jay ", "j")
InTXT$ = AF_Script(InTXT$, "kay ", "k")
InTXT$ = AF_Script(InTXT$, "el ", "l")
InTXT$ = AF_Script(InTXT$, "em ", "m")
InTXT$ = AF_Script(InTXT$, "en ", "n")
InTXT$ = AF_Script(InTXT$, "oh ", "o")
InTXT$ = AF_Script(InTXT$, "pee ", "p")
InTXT$ = AF_Script(InTXT$, "cue ", "q")
InTXT$ = AF_Script(InTXT$, "are ", "r")
InTXT$ = AF_Script(InTXT$, "ess ", "s")
InTXT$ = AF_Script(InTXT$, "tee ", "t")
InTXT$ = AF_Script(InTXT$, "you ", "u")
InTXT$ = AF_Script(InTXT$, "vee ", "v")
InTXT$ = AF_Script(InTXT$, "doubleyou ", "w")
InTXT$ = AF_Script(InTXT$, "ecks ", "x")
InTXT$ = AF_Script(InTXT$, "why ", "y")
InTXT$ = AF_Script(InTXT$, "zee ", "z")
InTXT$ = AF_Script(InTXT$, "ee ", "e")
T_SoundDecode$ = Trim$(InTXT$)
End Function
Function SoundEncode(ByVal InTXT As String) As String
Dim i
Dim l$
Dim T_SoundEncode$
Dim Final$
Dim n$
For i = 1 To Len(InTXT$)
    l$ = LCase(Mid$(InTXT$, i, 1))
    Select Case l$
        Case "a"
        n$ = "aigh"
        Case "b"
        n$ = "bee"
        Case "c"
        n$ = "see"
        Case "d"
        n$ = "dee"
        Case "e"
        n$ = "ee"
        Case "f"
        n$ = "eff"
        Case "g"
        n$ = "gee"
        Case "h"
        n$ = "aytch"
        Case "i"
        n$ = "eye"
        Case "j"
        n$ = "jay"
        Case "k"
        n$ = "kay"
        Case "l"
        n$ = "el"
        Case "m"
        n$ = "em"
        Case "n"
        n$ = "en"
        Case "o"
        n$ = "oh"
        Case "p"
        n$ = "pee"
        Case "q"
        n$ = "cue"
        Case "r"
        n$ = "are"
        Case "s"
        n$ = "ess"
        Case "t"
        n$ = "tee"
        Case "u"
        n$ = "you"
        Case "v"
        n$ = "vee"
        Case "w"
        n$ = "doubleyou"
        Case "x"
        n$ = "ecks"
        Case "y"
        n$ = "why"
        Case "z"
        n$ = "zee"
        Case Else
        n$ = l$
        End Select
    Final$ = Final$ & n$ & " "
    Next i
T_SoundEncode$ = Trim$(Final$)
End Function

Function StrFromLst(lstname As Listbox) As String
On Error GoTo daerror
Dim i
Dim dastr
For i = 1 To lstname.ListCount
    dastr = dastr & lstname.List(i - 1) & ","
Next i
StrFromLst = dastr
Exit Function
daerror:
StrFromLst = ""
End Function

Function AF_Script(ByVal Strt As String, ByVal ReplaceMe As String, ByVal ReplaceWith As String) As String
Dim x%
Dim Start$
Dim pos%
Start$ = Strt
Do While InStr(Start$, ReplaceMe$) <> 0
    x% = DoEvents()
    pos% = InStr(Start$, ReplaceMe$)
    Start$ = Left(Start$, pos% - 1) & ReplaceWith$ & Right(Start$, Len(Start$) - pos% - Len(ReplaceMe$) + 1)
    Loop
AF_Script$ = Start$
End Function

Public Sub addsentmailtocontrolwithdogbar(TheList As Control, progbar As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
             progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                 progbar.Value = index& - 1&
                 lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                 strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
                 lngtabkey& = InStr(strmailtext$, Chr$(9&))
                 lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                 strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub

Public Sub addsentmailtocontrol(TheList As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                 lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                 strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
                 lngtabkey& = InStr(strmailtext$, Chr$(9&))
                 lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                 strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub

Public Sub addpicturetoaolmdi(thepic As PictureBox)
    Dim lngaol As Long, lngmdi As Long, therect As RECT, lngrect As Long, lngchild As Long
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    
    lngrect& = GetWindowRect(lngmdi&, therect)
    Call StretchBlt(thepic.hDC, therect.Top, therect.Left, thepic.Height, thepic.Width, thepic.Picture, thepic.Top, thepic.Left, thepic.Picture.Width, thepic.Picture.Height, 0&)
    Call SetParent(thepic.hwnd, lngmdi&)
    Call BringWindowToTop(thepic.hwnd)
     lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Call BringWindowToTop(lngchild&)
End Sub


Public Sub addoldmailtocontrolwithdogbar(TheList As Control, progbar As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
             progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                 progbar.Value = index& - 1&
                 lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                 strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
                 lngtabkey& = InStr(strmailtext$, Chr$(9&))
                 lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                 strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub

Public Function ProfileGet(ScreenName As String) As String
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = GetText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
   sendchat "They dont have a profile =("
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    MsgBox "" & pString$ & "", 16, "Profile"
    End If
End Function

Public Sub addfavoritefolder(strfoldername As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long, lngmenu As Long, lngcheck As Long
    Dim lngcheck2 As Long, lngedit As Long, lngedit2 As Long, lngedit3 As Long
    Dim lngfavwin As Long, lngfavwin2 As Long, lngicon2 As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
         lngfavwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Favorite Places")
         lngicon& = FindWindowEx(lngfavwin&, 0&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngfavwin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngfavwin& <> 0& And lngicon& <> 0&
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
         lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
         lngcheck& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Checkbox", vbNullString)
         lngcheck2& = FindWindowEx(lngfavwin2&, lngcheck&, "_AOL_Checkbox", vbNullString)
         lngedit& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Edit", vbNullString)
         lngedit2& = FindWindowEx(lngfavwin2&, lngedit&, "_AOL_Edit", vbNullString)
         lngedit3& = FindWindowEx(lngfavwin2&, lngedit2&, "_AOL_Edit", vbNullString)
         lngicon& = FindWindowEx(lngfavwin2&, 0, "_AOL_Icon", vbNullString)
         lngicon2& = FindWindowEx(lngfavwin2&, lngicon&, "_AOL_Icon", vbNullString)
        Pause 0.2
    Loop Until lngfavwin2& <> 0& And lngcheck& <> 0& And lngcheck2& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngedit3& <> 0& And lngicon& <> 0& And lngicon2& <> 0&
    Call PostMessage(lngcheck2&, BM_SETCHECK, True, 0&)
    Call PostMessage(lngcheck2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngcheck2&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngcheck2&, BM_SETCHECK, True, 0&)
    Do: DoEvents
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfoldername$)
        Call PostMessage(lngicon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon2&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(lngicon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon2&, WM_LBUTTONUP, 0&, 0&)
         lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
    Loop Until lngfavwin2& = 0&
    Call windowclose(lngfavwin&)
End Sub
Function advEncrypt(Text As String, MainKey As Integer, Key1 As Integer, Key2 As Integer, Key3 As Integer, Key4 As Integer, Key5 As Integer, Key6 As Integer, Key7 As Integer, Key8 As Integer, Key9 As Integer) As String


    Dim a As Integer
    Dim i As Integer
    Dim endstr As String
    Dim endstr2 As String
    Dim endstr3 As String
    Dim KeyPos As Integer
    Dim Key(8) As Integer
    KeyPos = 1
    Key(0) = Key1
    Key(1) = Key2
    Key(2) = Key3
    Key(3) = Key4
    Key(4) = Key5
    Key(5) = Key6
    Key(6) = Key7
    Key(7) = Key8
    Key(8) = Key9
    If MainKey > 255 Then
        MainKey = 255
    ElseIf MainKey < 1 Then
        MainKey = 1
    End If


    If Key(0) > 100 Then
        Key(0) = 100
    ElseIf Key(0) < 1 Then
        Key(0) = 1
    End If


    If Key(1) > 100 Then
        Key(1) = 100
    ElseIf Key(1) < 1 Then
        Key(1) = 1
    End If


    If Key(2) > 100 Then
        Key(2) = 100
    ElseIf Key(2) < 1 Then
        Key(2) = 1
    End If


    If Key(3) > 100 Then
        Key(3) = 100
    ElseIf Key(3) < 1 Then
        Key(3) = 1
    End If


    If Key(4) > 100 Then
        Key(4) = 100
    ElseIf Key(4) < 1 Then
        Key(4) = 1
    End If


    If Key(5) > 100 Then
        Key(5) = 100
    ElseIf Key(5) < 1 Then
        Key(5) = 1
    End If


    If Key(6) > 100 Then
        Key(6) = 100
    ElseIf Key(6) < 1 Then
        Key(6) = 1
    End If


    If Key(7) > 100 Then
        Key(7) = 100
    ElseIf Key(7) < 1 Then
        Key(7) = 1
    End If


    If Key(8) > 100 Then
        Key(8) = 100
    ElseIf Key(8) < 1 Then
        Key(8) = 1
    End If


    For i = 1 To Len(Text$)
        If Asc(Mid$(Text$, i, 1)) + MainKey > 255 Then
            a = Asc(Mid$(Text$, i, 1)) + (MainKey - 255)
            endstr$ = endstr$ & Chr$(a)
        Else
            a = Asc(Mid$(Text$, i, 1)) + MainKey
            endstr$ = endstr$ & Chr$(a)
        End If


        MainKey = Int((((Key(0) + Key(1) - MainKey) / Key(2)) * (Key(3) - (Key(4) + Key(5)))) + ((Key(6) / Key(7)) + Key(8)))
        MainKey = Int(MainKey - Int((((Key(8) + Key(7)) / Key(6)) * (Key(5) - (Key(4) + Key(3)))) + ((Key(2) / Key(1)) + Key(0))))
CheckMainKey:
        If MainKey > 255 Then
            MainKey = MainKey - 255
            GoTo CheckMainKey
        End If


        Key(0) = Int(Key(0) + Int(Key(1) / 2))
        Key(1) = Int(Key(1) - Int(Key(2) * 2))
        Key(2) = Int(Key(2) + Int(Key(3) / 2))
        Key(3) = Int(Key(3) - Int(Key(4) * 2))
        Key(4) = Int(Key(4) + Int(Key(5) / 2))
        Key(5) = Int(Key(5) - Int(Key(6) * 2))
        Key(6) = Int(Key(6) + Int(Key(7) / 2))
        Key(7) = Int(Key(7) - Int(Key(8) * 2))
        Key(8) = Int(Key(8) + Int(Key(0) / 2))
        If Key(0) > 100 Then
            Key(0) = 100
        ElseIf Key(0) < 1 Then
            Key(0) = 1
        End If


        If Key(1) > 100 Then
            Key(1) = 100
        ElseIf Key(1) < 1 Then
            Key(1) = 1
        End If


        If Key(2) > 100 Then
            Key(2) = 100
        ElseIf Key(2) < 1 Then
            Key(2) = 1
        End If


        If Key(3) > 100 Then
            Key(3) = 100
        ElseIf Key(3) < 1 Then
            Key(3) = 1
        End If


        If Key(4) > 100 Then
            Key(4) = 100
        ElseIf Key(4) < 1 Then
            Key(4) = 1
        End If


        If Key(5) > 100 Then
            Key(5) = 100
        ElseIf Key(5) < 1 Then
            Key(5) = 1
        End If


        If Key(6) > 100 Then
            Key(6) = 100
        ElseIf Key(6) < 1 Then
            Key(6) = 1
        End If


        If Key(7) > 100 Then
            Key(7) = 100
        ElseIf Key(7) < 1 Then
            Key(7) = 1
        End If


        If Key(8) > 100 Then
            Key(8) = 100
        ElseIf Key(8) < 1 Then
            Key(8) = 1
        End If


    Next i


    advEncrypt = endstr$
End Function


Function advDecrypt(Text As String, MainKey As Integer, Key1 As Integer, Key2 As Integer, Key3 As Integer, Key4 As Integer, Key5 As Integer, Key6 As Integer, Key7 As Integer, Key8 As Integer, Key9 As Integer) As String


    Dim a As Integer
    Dim i As Integer
    Dim endstr As String
    Dim endstr2 As String
    Dim endstr3 As String
    Dim KeyPos As Integer
    Dim Key(8) As Integer
    KeyPos = 1
    Key(0) = Key1
    Key(1) = Key2
    Key(2) = Key3
    Key(3) = Key4
    Key(4) = Key5
    Key(5) = Key6
    Key(6) = Key7
    Key(7) = Key8
    Key(8) = Key9
    If MainKey > 255 Then
        MainKey = 255
    ElseIf MainKey < 1 Then
        MainKey = 1
    End If


    If Key(0) > 100 Then
        Key(0) = 100
    ElseIf Key(0) < 1 Then
        Key(0) = 1
    End If


    If Key(1) > 100 Then
        Key(1) = 100
    ElseIf Key(1) < 1 Then
        Key(1) = 1
    End If


    If Key(2) > 100 Then
        Key(2) = 100
    ElseIf Key(2) < 1 Then
        Key(2) = 1
    End If


    If Key(3) > 100 Then
        Key(3) = 100
    ElseIf Key(3) < 1 Then
        Key(3) = 1
    End If


    If Key(4) > 100 Then
        Key(4) = 100
    ElseIf Key(4) < 1 Then
        Key(4) = 1
    End If


    If Key(5) > 100 Then
        Key(5) = 100
    ElseIf Key(5) < 1 Then
        Key(5) = 1
    End If


    If Key(6) > 100 Then
        Key(6) = 100
    ElseIf Key(6) < 1 Then
        Key(6) = 1
    End If


    If Key(7) > 100 Then
        Key(7) = 100
    ElseIf Key(7) < 1 Then
        Key(7) = 1
    End If


    If Key(8) > 100 Then
        Key(8) = 100
    ElseIf Key(8) < 1 Then
        Key(8) = 1
    End If


    For i = 1 To Len(Text$)
        If Asc(Mid$(Text$, i, 1)) - MainKey < 1 Then
            a = (Asc(Mid$(Text$, i, 1)) - MainKey) + 255
            endstr$ = endstr$ & Chr$(a)
        Else
            a = Asc(Mid$(Text$, i, 1)) - MainKey
            endstr$ = endstr$ & Chr$(a)
        End If


        MainKey = Int((((Key(0) + Key(1) - MainKey) / Key(2)) * (Key(3) - (Key(4) + Key(5)))) + ((Key(6) / Key(7)) + Key(8)))
        MainKey = Int(MainKey - Int((((Key(8) + Key(7)) / Key(6)) * (Key(5) - (Key(4) + Key(3)))) + ((Key(2) / Key(1)) + Key(0))))
CheckMainKey:
        If MainKey > 255 Then
            MainKey = MainKey - 255
            GoTo CheckMainKey
        End If


        Key(0) = Int(Key(0) + Int(Key(1) / 2))
        Key(1) = Int(Key(1) - Int(Key(2) * 2))
        Key(2) = Int(Key(2) + Int(Key(3) / 2))
        Key(3) = Int(Key(3) - Int(Key(4) * 2))
        Key(4) = Int(Key(4) + Int(Key(5) / 2))
        Key(5) = Int(Key(5) - Int(Key(6) * 2))
        Key(6) = Int(Key(6) + Int(Key(7) / 2))
        Key(7) = Int(Key(7) - Int(Key(8) * 2))
        Key(8) = Int(Key(8) + Int(Key(0) / 2))
        If Key(0) > 100 Then
            Key(0) = 100
        ElseIf Key(0) < 1 Then
            Key(0) = 1
        End If


        If Key(1) > 100 Then
            Key(1) = 100
        ElseIf Key(1) < 1 Then
            Key(1) = 1
        End If


        If Key(2) > 100 Then
            Key(2) = 100
        ElseIf Key(2) < 1 Then
            Key(2) = 1
        End If


        If Key(3) > 100 Then
            Key(3) = 100
        ElseIf Key(3) < 1 Then
            Key(3) = 1
        End If


        If Key(4) > 100 Then
            Key(4) = 100
        ElseIf Key(4) < 1 Then
            Key(4) = 1
        End If


        If Key(5) > 100 Then
            Key(5) = 100
        ElseIf Key(5) < 1 Then
            Key(5) = 1
        End If


        If Key(6) > 100 Then
            Key(6) = 100
        ElseIf Key(6) < 1 Then
            Key(6) = 1
        End If


        If Key(7) > 100 Then
            Key(7) = 100
        ElseIf Key(7) < 1 Then
            Key(7) = 1
        End If


        If Key(8) > 100 Then
            Key(8) = 100
        ElseIf Key(8) < 1 Then
            Key(8) = 1
        End If


    Next i


    advDecrypt = endstr$
End Function


Sub DeleteItem(lst As Listbox, Item$)
Dim a
On Error Resume Next
Do
DoEvents:
If LCase$(lst.List(a)) = LCase$(Item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Function AddListToString(TheList As Listbox)
Dim DoList
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function

'
Public Function CheckIMs(Person As String) As Boolean
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
sendchat "" & Person & " -" & Chr(13) & "im's on"

Else
     If InStr(oString$, "is not currently signed on.") <> 0 Then
        CheckIMs = False
sendchat "" & Person & " -" & Chr(13) & "is not signed on or is Ghosting."
Else
If InStr(oString$, "cannot currently receive Instant Messages.") <> 0 Then
        CheckIMs = False
sendchat "" & Person & " -" & Chr(13) & "Im's are OFF."
End If
End If
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
End Function


Public Sub addfavoriteitem(strdescription As String, strurl As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long, lngmenu As Long, lngcheck As Long
    Dim lngcheck2 As Long, lngedit As Long, lngedit2 As Long, lngedit3 As Long
    Dim lngfavwin As Long, lngfavwin2 As Long, lngicon2 As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
         lngfavwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Favorite Places")
         lngicon& = FindWindowEx(lngfavwin&, 0&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngfavwin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngfavwin& <> 0& And lngicon& <> 0&
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
        lngcheck& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Checkbox", vbNullString)
        lngcheck2& = FindWindowEx(lngfavwin2&, lngcheck&, "_AOL_Checkbox", vbNullString)
        lngedit& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Edit", vbNullString)
        lngedit2& = FindWindowEx(lngfavwin2&, lngedit&, "_AOL_Edit", vbNullString)
        lngedit3& = FindWindowEx(lngfavwin2&, lngedit2&, "_AOL_Edit", vbNullString)
        lngicon& = FindWindowEx(lngfavwin2&, 0, "_AOL_Icon", vbNullString)
        lngicon2& = FindWindowEx(lngfavwin2&, lngicon&, "_AOL_Icon", vbNullString)
        Pause 0.2
    Loop Until lngfavwin2& <> 0& And lngcheck& <> 0& And lngcheck2& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngedit3& <> 0& And lngicon& <> 0& And lngicon2& <> 0&
    Call SendMessageByString(lngedit2&, WM_SETTEXT, 0&, strdescription$)
    Call SendMessageByString(lngedit3&, WM_SETTEXT, 0&, strurl$)
    Call PostMessage(lngicon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
    Loop Until lngfavwin2& = 0&
    Call windowclose(lngfavwin&)
End Sub

Public Sub MailForward(sendto As String, Message As String, DeleteFwd As Boolean)
    Dim AOL As Long, MDI As Long, ErRor As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String
    Do
        DoEvents
        OpenForward& = FindForwardWindow
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    If DeleteFwd = True Then
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
        DoEvents
    End If
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, sendto$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    DoEvents
    Do Until OpenSend& = 0& Or ErRor& <> 0&
        DoEvents
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
        ErRor& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        timeout (1)
    Loop
    If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
End Sub
Sub ClickForward()
Dim hWnd1%
Dim hWnd2%
Dim hWnd3%
Dim hWnd4%

hWnd1% = FindWindow("AOL Frame25", vbNullString)
hWnd2% = FindWindowEx(hWnd1%, 0&, "MDIClient", vbNullString)
hWnd3% = FindWindowEx(hWnd2%, 0&, "AOL Child", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, 0&, "_AOL_Icon", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, hWnd4%, "_AOL_Icon", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, hWnd4%, "_AOL_Icon", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, hWnd4%, "_AOL_Icon", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, hWnd4%, "_AOL_Icon", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, hWnd4%, "_AOL_Icon", vbNullString)
hWnd4% = FindWindowEx(hWnd3%, hWnd4%, "_AOL_Icon", vbNullString)

ClickIcon (hWnd4%)
End Sub
Public Sub AddRoomToListbox(TheList As Listbox, adduser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or adduser = True Then
                TheList.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub AddRoomToListbox7(TheList As Listbox, adduser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 28
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            
            If ScreenName$ <> GetUser$ Or adduser = True Then
                TheList.AddItem ScreenName$
            End If

        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub signon(ScreenName As String, password As String)


    Dim AOL As Long, MDI As Long, signonwin As Long, namelist As Long
    Dim passwordbox As Long, SignOnButton As Long, snindex As Long
    Dim index As Long, indexlength As Long, thestring As String
    Dim rList As Long, sThread As Long, mThread As Long, itmHold As Long
    Dim psnHold As Long, rBytes As Long, cProcess As Long, combocount As Long
    Dim guestwin As Long, guestsn As Long, guestpw As Long, guestok As Long
    Dim child As Long, signonbutton1 As Long, errorwin As Long, OKButton As Long
    Dim guestcancel As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    signonwin& = findsignonwin()
    If signonwin& = 0& Then
        Do: DoEvents
            child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
            Call SendMessage(child&, WM_CLOSE, 0&, 0&)
        Loop Until child& = 0&


        Call RunMenu(3&, 0&)
    End If


    Do: DoEvents
        signonwin& = findsignonwin()
        namelist& = FindWindowEx(signonwin&, 0&, "_AOL_Combobox", vbNullString)
        passwordbox& = FindWindowEx(signonwin&, 0&, "_AOL_Edit", vbNullString)
        signonbutton1& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
        signonbutton1& = FindWindowEx(signonwin&, signonbutton1&, "_AOL_Icon", vbNullString)
        signonbutton1& = FindWindowEx(signonwin&, signonbutton1&, "_AOL_Icon", vbNullString)
        SignOnButton& = FindWindowEx(signonwin&, signonbutton1&, "_AOL_Icon", vbNullString)
    Loop Until signonwin& <> 0& And namelist& <> 0& And passwordbox& <> 0& And SignOnButton& <> 0&


    combocount& = SendMessage(namelist&, CB_GETCOUNT, 0&, 0&)
    Call PostMessage(namelist&, CB_SETCURSEL, combocount& - 1, 0&)
    Call PostMessage(SignOnButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(SignOnButton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        guestwin& = FindWindow("_AOL_Modal", vbNullString)
        guestsn& = FindWindowEx(guestwin&, 0&, "_AOL_Edit", vbNullString)
        guestpw& = FindWindowEx(guestwin&, guestsn&, "_AOL_Edit", vbNullString)
        guestok& = FindWindowEx(guestwin&, 0&, "_AOL_Icon", vbNullString)
        guestcancel& = FindWindowEx(guestwin&, guestok&, "_AOL_Icon", vbNullString)
    Loop Until guestwin& <> 0& And guestsn& <> 0& And guestpw& <> 0& And guestok& <> 0&


    Call SendMessageByString(guestsn&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(guestpw&, WM_SETTEXT, 0&, password$)
    Call PostMessage(guestok&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(guestok&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        guestwin& = FindWindow("_AOL_Modal", vbNullString)
        errorwin& = FindWindow("#32770", "America Online")
    Loop Until guestwin& = 0& Or errorwin& <> 0&


    If errorwin& <> 0& Then
        OKButton& = FindWindowEx(errorwin&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
        Call PostMessage(guestcancel&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(guestcancel&, WM_LBUTTONUP, 0&, 0&)
    End If


End Sub



Public Function findsignonwin() As Long


    Dim AOL As Long, MDI As Long, child As Long, Combo1 As Long
    Dim combo2 As Long, static1 As Long, static2 As Long, static3 As Long
    Dim static4 As Long, edit1 As Long, icon1 As Long, icon2 As Long
    Dim icon3 As Long, icon4 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Combo1& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
    combo2& = FindWindowEx(child&, Combo1&, "_AOL_Combobox", vbNullString)
    static1& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    static2& = FindWindowEx(child&, static1&, "_AOL_Static", vbNullString)
    static3& = FindWindowEx(child&, static2&, "_AOL_Static", vbNullString)
    static4& = FindWindowEx(child&, static3&, "_AOL_Static", vbNullString)
    edit1& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
    icon1& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    icon2& = FindWindowEx(child&, icon1&, "_AOL_Icon", vbNullString)
    icon3& = FindWindowEx(child&, icon2&, "_AOL_Icon", vbNullString)
    icon4& = FindWindowEx(child&, icon3&, "_AOL_Icon", vbNullString)
    If Combo1& <> 0& And combo2 <> 0& And static1& <> 0& And static2& <> 0& And static3& <> 0& And static4& <> 0& And edit1& <> 0& And icon1& <> 0& And icon2& <> 0& And icon3& <> 0& And icon4& <> 0& Then
        findsignonwin& = child&
    Else
        Do: DoEvents
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Combo1& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
            combo2& = FindWindowEx(child&, Combo1&, "_AOL_Combobox", vbNullString)
            static1& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            static2& = FindWindowEx(child&, static1&, "_AOL_Static", vbNullString)
            static3& = FindWindowEx(child&, static2&, "_AOL_Static", vbNullString)
            static4& = FindWindowEx(child&, static3&, "_AOL_Static", vbNullString)
            edit1& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
            icon1& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            icon2& = FindWindowEx(child&, icon1&, "_AOL_Icon", vbNullString)
            icon3& = FindWindowEx(child&, icon2&, "_AOL_Icon", vbNullString)
            icon4& = FindWindowEx(child&, icon3&, "_AOL_Icon", vbNullString)
            If Combo1& <> 0& And combo2 <> 0& And static1& <> 0& And static2& <> 0& And static3& <> 0& And static4& <> 0& And edit1& <> 0& And icon1& <> 0& And icon2& <> 0& And icon3& <> 0& And icon4& <> 0& Then
                findsignonwin& = child&
                Exit Function
            End If


        Loop Until child& = 0&


        Exit Function
    End If


    findsignonwin& = child&
End Function


Public Sub RunMenu(TopMenu As Long, SubMenu As Long)


    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub
Public Function CheckAlive(ScreenName As String) As Boolean
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call SendMail("*, " & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
sendchat "</b><Font Face=Arial>‹- " & ScreenName$ & " ­ dead"
    Else
        CheckAlive = True
sendchat "</b><Font Face=Arial>‹- " & ScreenName$ & " ­ active"
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        NoWindow& = FindWindow("#32770", "America Online")
        NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
    Loop Until NoWindow& <> 0& And NoButton& <> 0
    Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function

Function StartPlay()


    mciSendString "play cd", 0, 0, 0
End Function


Function SetTrack(track%)


    mciSendString "seek cd to " & Str(track), 0, 0, 0
End Function


Function StopPlay()


    mciSendString "stop cd wait", 0, 0, 0
End Function


Function PausePlay()


    mciSendString "pause cd", 0, 0, 0
End Function







Function SetFormat_tmsf()


    mciSendString "set cd time format tmsf wait", 0, 0, 0
End Function


Function SetFormat_milliseconds()


    mciSendString "set cd time format milliseconds", 0, 0, 0
End Function


Function CheckCD%()


    Dim s As String * 30
    mciSendString "status cd media present", s, Len(s), 0
    CheckCD = s
End Function


Function GetNumTracks%()


    Dim s As String * 30
    mciSendString "status cd number of tracks wait", s, Len(s), 0
    GetNumTracks = CInt(Mid$(s, 1, 2))
End Function





Function GetTrackLength$(TrackNum%)


    Dim s As String * 30
    mciSendString "status cd length track " & TrackNum, s, Len(s), 0
    GetTrackLength = s
End Function


Sub GetCDPosition(track%, min%, SEC%)


    Dim s As String * 30
    mciSendString "status cd position", s, Len(s), 0
    track = CInt(Mid$(s, 1, 2))
    min = CInt(Mid$(s, 4, 2))
    SEC = CInt(Mid$(s, 7, 2))
End Sub


Function CheckIfPlaying%()


    CheckIfPlaying = 0
    Dim s As String * 30
    mciSendString "status cd mode", s, Len(s), 0
    If Mid$(s, 1, 7) = "playing" Then CheckIfPlaying = 1
End Function


'|---------------------|
'|---Automated Tasks---|
'V---------------------V
Function SeekCDtoX(track%)


    StopPlay
    SetTrack track
    StartPlay
End Function


Function ReadyDevice()


    UnloadAll
    SetCDPlayerReady
    SetFormat_tmsf
End Function


Function FastForward(Spd%)


    Dim s As String * 40
    SetFormat_milliseconds
    mciSendString "status cd position wait", s, Len(s), 0
    CheckIfPlaying%
    If CheckIfPlaying = 1 Then
        mciSendString "play cd from " & CStr(CLng(s) + Spd), 0, 0, 0
    Else
        mciSendString "seek cd to " & CStr(CLng(s) + Spd), 0, 0, 0
    End If


    SetFormat_tmsf
End Function


Function ReWind(Spd%)


    Dim s As String * 40
    SetFormat_milliseconds
    mciSendString "status cd position wait", s, Len(s), 0
    CheckIfPlaying%
    If CheckIfPlaying = 1 Then
        mciSendString "play cd from " & CStr(CLng(s) - Spd), 0, 0, 0
    Else
        mciSendString "seek cd to " & CStr(CLng(s) - Spd), 0, 0, 0
    End If


    SetFormat_tmsf
End Function
Public Sub addflashmailtocontrol(TheList As Control)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    Dim strstring As String, index As Long, lnglen As Long, thetab As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
     lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
     strstring$ = String$(255&, 0&)
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
        For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&: DoEvents
             lnglen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
             strstring$ = String(lnglen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETtext, index&, strstring$)
             thetab& = InStr(strstring$, Chr$(9&))
             thetab& = InStr(thetab& + 1&, strstring$, Chr$(9&))
             strstring$ = Right$(strstring$, Len(strstring$) - thetab&)
             strstring$ = ReplaceString(strstring$, Chr$(0&), "")
            TheList.AddItem strstring$
        Next index&
    End If
End Sub


Public Function IMSender() As String
    Dim im As Long, Caption As String
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function


Public Function ErrorName(Name As Long) As String
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long, TempString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = GetText(ErrorTextWindow&)
    NameCount& = lineCount(ErrorString$) - 2
    If NameCount& < Name& Then Exit Function
    TempString$ = LineFromString(ErrorString$, Name& + 2)
    TempString$ = Left(TempString$, InStr(TempString$, "-") - 2)
    ErrorName$ = TempString$
End Function


Public Function LineFromString(MyString As String, Line As Long) As String
    Dim theline As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = lineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function

Function lineCount(myInFile As String) As Long
    Dim lFileSize As Long, lChunk As Long
    Dim bFile() As Byte
    Dim lSize As Long
    Dim strText As String
    
    'the size of the chunk to read in. You c
    '     an experiment
    'with this to see what works fastest.
    lSize = CLng(1024) * 10
    
    'size the array to the chunk size
    ReDim bFile(lSize - 1) As Byte
    
    Open myInFile For Binary As #1
    'get the file size
    lFileSize = LOF(1)
    
    'set the chunk number to 1
    lChunk = 1


    Do While (lSize * lChunk) < lFileSize
        'get the data from the in file
        Get #1, , bFile
        strText = StrConv(bFile, vbUnicode)
        'get the line count for this chunk
        lineCount = lineCount + searchText(strText)
        'increment the chunk count
        lChunk = lChunk + 1
    Loop
    
    'redim the array to the remaining size
    ReDim bFile((lFileSize - (lSize * (lChunk - 1))) - 1) As Byte
    'get the remaining data
    Get #1, , bFile
    strText = StrConv(bFile, vbUnicode)
    'get line count for this chunk
    lineCount = lineCount + searchText(strText)
    
    'close the file
    Close #1
    
    lineCount = lineCount + 1
End Function


Public Function ErrorNameCount() As Long
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = GetText(ErrorTextWindow&)
    NameCount& = lineCount(ErrorString$) - 2
    ErrorNameCount& = NameCount&
End Function



Function FindEjectTab()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
FindEjectTab = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Function FindEjectTab9()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
FindEjectTab9 = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
End Function
Function FindPeopleHereTab()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
FindPeopleHereTab = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Function FindPeopleHereTab9()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
FindPeopleHereTab9 = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)

End Function
Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim ArrayCount As Integer
Dim Document
Dim FState
    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deed, then return that index.
    For i = 1 To ArrayCount
        If FState(i).Deed Then
            FindFreeIndex = i
            FState(i).Deed = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deed, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve FState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function





Sub WriteRecentFiles(OPENFILENAME)
    Dim ThisApp
    Dim ThisKey
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim i, J As Integer
    Dim strFile, Key As String

    ' Copy RecentFile1 to RecentFile2, and so on.
    For i = 3 To 1 Step -1
        Key = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, Key)
        If strFile <> "" Then
            Key = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, Key, strFile
        End If
    Next i
  
    ' Write the open file to first recent file.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OPENFILENAME
End Sub


Public Sub addflashmailtocontrolwithdogbar(TheList As Control, progbar As Control)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    Dim strstring As String, index As Long, lnglen As Long, thetab As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
     lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
     strstring$ = String$(255&, 0&)
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
         progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&: DoEvents
             progbar.Value = index& - 1&
             lnglen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
             strstring$ = String(lnglen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETtext, index&, strstring$)
             thetab& = InStr(strstring$, Chr$(9&))
             thetab& = InStr(thetab& + 1&, strstring$, Chr$(9&))
             strstring$ = Right$(strstring$, Len(strstring$) - thetab&)
             strstring$ = ReplaceString(strstring$, Chr$(0&), "")
            TheList.AddItem strstring$
        Next index&
    End If
End Sub

Public Sub MailOpenEmailNew(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub


Public Sub MailToListNew(TheList As Listbox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETtext, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem AddMails& + 1 & "‡    " & MyString$
    Next AddMails&

End Sub
Public Sub MailOpenNew()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
End Sub

Public Function FindSendWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim SendStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindSendWindow& = 0&
End Function
Public Sub addfontstocontrol(TheList As Control)
    Dim lngindex As Long, lngsw2 As Long, lngsw As Long
    TheList.Clear
    For lngindex& = 1& To Screen.FontCount
        DoEvents
        TheList.AddItem (Screen.Fonts(lngindex& - 1&))
    Next lngindex&
End Sub

Public Sub addnewmailtocontrol(TheList As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                 lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                 strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
                 lngtabkey& = InStr(strmailtext$, Chr$(9&))
                 lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                 strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub

Public Sub addnewmailtocontrolwithdogbar(TheList As Control, progbar As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
             progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                 progbar.Value = index& - 1&
                 lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                 strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
                 lngtabkey& = InStr(strmailtext$, Chr$(9&))
                 lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                 strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub

Public Sub phoneannoy(hehenum, PhoneNumber)
Dim thenum
Do Until thenum = hehenum
DoEvents:
Open "COM1:" For Output As #1 'or COM2
Print #1, "ATDT" & PhoneNumber & Chr$(13)
thenum = thenum + 1
Close #1
Loop
End Sub

Public Function doeswindowexist(lngwindow As Long) As Boolean
    If IsWindow(lngwindow&) Then
         doeswindowexist = True
    Else
         doeswindowexist = False
    End If
End Function

Sub ClickSendAfterError(Recipiants)
Dim OpenForward&
Dim OpenSend&
Dim EditTo&
Dim fCombo&
Dim Rich&
Dim Combo&
Dim Button2&
Dim SendButton&
Dim Button1&
Dim DoIt&
Dim EditSubject&
Dim EditCC&
Pause 0.8
        DoEvents
        OpenForward& = FindForwardWindow
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Recipiants)
       
       OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
   




End Sub

Public Sub enableaoltoolbar()
    'this will enable the aol toolbar and items when signed off
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Call EnableWindow(lngicon&, True)
    For index& = 0& To 20&
         lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
        Call EnableWindow(lngicon&, True)
    Next index&
     lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
End Sub

Public Sub fadeformdiagonallyfast(frmname As Form, color1 As Long, color2 As Long)
    'call this in the form_paint of your form
    'this will fade it from top to bottom
    'preset colors are in the dec's of this bas
    
    'works with most size forms
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
     lngblue1& = Int(color1& / 65536)
     lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
     lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
     lngblue2& = Int(color2& / 65536)
     lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
     lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
     frmname.ScaleMode = vbPixels
     frmname.ScaleWidth = 256&
     frmname.DrawWidth = 2&
     frmname.DrawMode = vbCopyPen
     frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (index&, 0&)-(Screen.Height, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
        frmname.Line (0&, index&)-(Screen.Width, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
End Sub

Public Sub fadeformdiagonallyslow(frmname As Form, color1 As Long, color2 As Long)
    'call this in the form_paint of your form
    'this will fade it diagonally
    'preset colors are in the dec's of this bas
    
    'works with most sized forms
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
     lngblue1& = Int(color1& / 65536)
     lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
     lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
     lngblue2& = Int(color2& / 65536)
     lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
     lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
     frmname.ScaleMode = vbPixels
     frmname.ScaleWidth = 256&
     frmname.DrawWidth = 2&
     frmname.DrawMode = vbCopyPen
     frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (index&, 0&)-(Screen.Height, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
        frmname.Line (0&, index&)-(index& + 1&, Screen.Width), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
End Sub

Public Sub fadeformdiagonallywithborder(frmname As Form, color1 As Long, color2 As Long)
    'call this in the form_paint of your form
    'this will fade it from top to bottom
    'preset colors are in the dec's of this bas
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
     lngblue1& = Int(color1& / 65536)
     lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
     lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
     lngblue2& = Int(color2& / 65536)
     lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
     lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
     frmname.ScaleMode = vbPixels
     frmname.ScaleWidth = 256&
     frmname.DrawWidth = 2&
     frmname.DrawMode = vbCopyPen
     frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (index&, 0&)-(5, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
        frmname.Line (0&, index&)-(Screen.Width, 5), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
End Sub

Public Sub fadeobjectcircle(theobject As Object, Blue As Boolean, Red As Boolean, Green As Boolean)
    'IDEA from fireball
    Dim lngwidth As Long, lngheight As Long, lngblueval As Long, lngobjwidth As Long, lngredval As Long, lnggreenval As Long
     theobject.FillStyle = 0&
     lngwidth& = theobject.Width
     lngheight& = theobject.Height
    If Blue = True Then
        Do Until lngblueval& = 255&: DoEvents
             lngblueval& = lngblueval& + 1&
             lngwidth& = lngwidth& - theobject.Width / 255&
             theobject.FillColor = RGB(0&, 0&, lngblueval&)
            If lngwidth& < 0& Then
                Exit Sub
            Else
                theobject.Circle (theobject.Width / 2&, theobject.Height / 2&), lngwidth&, RGB(0&, 0&, lngblueval&)
            End If
        Loop
    ElseIf Red = True Then
        Do Until lngredval& = 255&: DoEvents
             lngredval& = lngredval& + 1&
             lngwidth& = lngwidth& - theobject.Width / 255&
             theobject.FillColor = RGB(lngredval&, 0&, 0&)
            If lngwidth& < 0& Then
                Exit Sub
            Else
                theobject.Circle (theobject.Width / 2&, theobject.Height / 2&), lngwidth&, RGB(lngredval&, 0&, 0&)
            End If
        Loop
    ElseIf Green = True Then
        Do Until lnggreenval& = 255&: DoEvents
             lnggreenval& = lnggreenval& + 1&
             lngwidth& = lngwidth& - theobject.Width / 255&
             theobject.FillColor = RGB(0&, lnggreenval&, 0&)
            If lngwidth& < 0& Then
                Exit Sub
            Else
                theobject.Circle (theobject.Width / 2&, theobject.Height / 2&), lngwidth&, RGB(0&, lnggreenval&, 0&)
            End If
        Loop
    End If
End Sub

Public Sub fadeobjectside2side2colors(theobject As Object, color1 As Long, color2 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 2&, 3&) As Integer
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
     lngcolors(1&, 1&) = lngred1&
     lngcolors(1&, 2&) = lnggreen1&
     lngcolors(1&, 3&) = lngblue1&
     lngcolors(2&, 1&) = lngred2&
     lngcolors(2&, 2&) = lnggreen2&
     lngcolors(2&, 3&) = lngblue2&
    ReDim lngcolors2(1& To 2&, 1& To 3&) As Double
     obj100th = theobject.ScaleWidth / 100&
     lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 2&
        For index2& = 1& To 3&
             lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
     theobject.BackColor = RGB(lngcolors2(2&, 1&), lngcolors2(2&, 2&), lngcolors2(2&, 3&))
    For index& = 1& To (2& - 1&)
         lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (2& - 1&))
         lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (2& - 1&))
         lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (2& - 1&))
        For index2& = 1& To (100& / (2& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
             lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
             lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
             lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
             lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub

Public Sub fadeobjectside2side3colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 3&, 3&) As Integer
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strcolor3$ = gethexfromrgb(color3&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     strred3$ = "&h" & Right$(strcolor3$, 2&)
     strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
     strblue3$ = "&h" & Left$(strcolor3$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
     lngred3& = Val(strred3$)
     lnggreen3& = Val(strgreen3$)
     lngblue3& = Val(strblue3$)
     lngcolors(1&, 1&) = lngred1&
     lngcolors(1&, 2&) = lnggreen1&
     lngcolors(1&, 3&) = lngblue1&
     lngcolors(2&, 1&) = lngred2&
     lngcolors(2&, 2&) = lnggreen2&
     lngcolors(2&, 3&) = lngblue2&
     lngcolors(3&, 1&) = lngred3&
     lngcolors(3&, 2&) = lnggreen3&
     lngcolors(3&, 3&) = lngblue3&
    ReDim lngcolors2(1& To 3&, 1& To 3&) As Double
     obj100th = theobject.ScaleWidth / 100&
     lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 3&
        For index2& = 1& To 3&
             lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
     theobject.BackColor = RGB(lngcolors2(3&, 1&), lngcolors2(3&, 2&), lngcolors2(3&, 3&))
    For index& = 1& To (3& - 1&)
         lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (3& - 1&))
         lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (3& - 1&))
         lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (3& - 1&))
        For index2& = 1& To (100& / (3& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
             lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
             lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
             lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
             lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub

Public Sub fadeobjectside2side4colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long, color4 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    Dim strcolor4 As String, strred4 As String, strblue4 As String, strgreen4 As String
    Dim lngred4 As Long, lngblue4 As Long, lnggreen4 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 4&, 3&) As Integer
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strcolor3$ = gethexfromrgb(color3&)
     strcolor4$ = gethexfromrgb(color4&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     strred3$ = "&h" & Right$(strcolor3$, 2&)
     strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
     strblue3$ = "&h" & Left$(strcolor3$, 2&)
     strred4$ = "&h" & Right$(strcolor4$, 2&)
     strgreen4$ = "&h" & Mid$(strcolor4$, 3&, 2&)
     strblue4$ = "&h" & Left$(strcolor4$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
     lngred3& = Val(strred3$)
     lnggreen3& = Val(strgreen3$)
     lngblue3& = Val(strblue3$)
     lngred4& = Val(strred4$)
     lnggreen4& = Val(strgreen4$)
     lngblue4& = Val(strblue4$)
     lngcolors(1&, 1&) = lngred1&
     lngcolors(1&, 2&) = lnggreen1&
     lngcolors(1&, 3&) = lngblue1&
     lngcolors(2&, 1&) = lngred2&
     lngcolors(2&, 2&) = lnggreen2&
     lngcolors(2&, 3&) = lngblue2&
     lngcolors(3&, 1&) = lngred3&
     lngcolors(3&, 2&) = lnggreen3&
     lngcolors(3&, 3&) = lngblue3&
     lngcolors(4&, 1&) = lngred4&
     lngcolors(4&, 2&) = lnggreen4&
     lngcolors(4&, 3&) = lngblue4&
    ReDim lngcolors2(1& To 4&, 1& To 3&) As Double
     obj100th = theobject.ScaleWidth / 100&
     lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 4&
        For index2& = 1& To 3&
             lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
     theobject.BackColor = RGB(lngcolors2(4&, 1&), lngcolors2(4&, 2&), lngcolors2(4&, 3&))
    For index& = 1& To (4& - 1&)
         lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (4& - 1&))
         lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (4& - 1&))
         lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (4& - 1&))
        For index2& = 1& To (100& / (4& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
             lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
             lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
             lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
             lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub
Public Sub fadeobjecttop2bottom2colors(theobject As Object, color1 As Long, color2 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 2&, 3&) As Integer
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
     lngcolors(1&, 1&) = lngred1&
     lngcolors(1&, 2&) = lnggreen1&
     lngcolors(1&, 3&) = lngblue1&
     lngcolors(2&, 1&) = lngred2&
     lngcolors(2&, 2&) = lnggreen2&
     lngcolors(2&, 3&) = lngblue2&
    ReDim lngcolors2(1& To 2&, 1& To 3&) As Double
     obj100th = theobject.ScaleHeight / 100&
     lngswidth& = theobject.ScaleWidth
    theobject.Cls
    For index& = 1& To 2&
        For index2& = 1& To 3&
             lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
     theobject.BackColor = RGB(lngcolors2(2&, 1&), lngcolors2(2&, 2&), lngcolors2(2&, 3&))
    For index& = 1& To (2& - 1&)
         lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (2& - 1&))
         lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (2& - 1&))
         lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (2& - 1&))
        For index2& = 1 To (100& / (2& - 1&))
            theobject.Line (0&, lnghold)-(lngswidth&, lnghold + obj100th), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
             lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
             lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
             lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
             lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub

Public Sub fadeobjecttop2bottom3colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 3&, 3&) As Integer
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strcolor3$ = gethexfromrgb(color3&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     strred3$ = "&h" & Right$(strcolor3$, 2&)
     strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
     strblue3$ = "&h" & Left$(strcolor3$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
     lngred3& = Val(strred3$)
     lnggreen3& = Val(strgreen3$)
     lngblue3& = Val(strblue3$)
     lngcolors(1&, 1&) = lngred1&
     lngcolors(1&, 2&) = lnggreen1&
     lngcolors(1&, 3&) = lngblue1&
     lngcolors(2&, 1&) = lngred2&
     lngcolors(2&, 2&) = lnggreen2&
     lngcolors(2&, 3&) = lngblue2&
     lngcolors(3&, 1&) = lngred3&
     lngcolors(3&, 2&) = lnggreen3&
     lngcolors(3&, 3&) = lngblue3&
    ReDim lngcolors2(1& To 3&, 1& To 3&) As Double
     obj100th = theobject.ScaleHeight / 100&
     lngswidth& = theobject.ScaleWidth
    theobject.Cls
    For index& = 1& To 3&
        For index2& = 1& To 3&
             lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
     theobject.BackColor = RGB(lngcolors2(3&, 1&), lngcolors2(3&, 2&), lngcolors2(3&, 3&))
    For index& = 1& To (3& - 1&)
         lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (3& - 1&))
         lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (3& - 1&))
         lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (3& - 1&))
        For index2& = 1& To (100& / (3& - 1&))
            theobject.Line (0&, lnghold)-(lngswidth&, lnghold + obj100th), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
             lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
             lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
             lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
             lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub


Public Sub fadeobjecttop2bottom4colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long, color4 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    Dim strcolor4 As String, strred4 As String, strblue4 As String, strgreen4 As String
    Dim lngred4 As Long, lngblue4 As Long, lnggreen4 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 4&, 3&) As Integer
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strcolor3$ = gethexfromrgb(color3&)
     strcolor4$ = gethexfromrgb(color4&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     strred3$ = "&h" & Right$(strcolor3$, 2&)
     strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
     strblue3$ = "&h" & Left$(strcolor3$, 2&)
     strred4$ = "&h" & Right$(strcolor4$, 2&)
     strgreen4$ = "&h" & Mid$(strcolor4$, 3&, 2&)
     strblue4$ = "&h" & Left$(strcolor4$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
     lngred3& = Val(strred3$)
     lnggreen3& = Val(strgreen3$)
     lngblue3& = Val(strblue3$)
     lngred4& = Val(strred4$)
     lnggreen4& = Val(strgreen4$)
     lngblue4& = Val(strblue4$)
     lngcolors(1&, 1&) = lngred1&
     lngcolors(1&, 2&) = lnggreen1&
     lngcolors(1&, 3&) = lngblue1&
     lngcolors(2&, 1&) = lngred2&
     lngcolors(2&, 2&) = lnggreen2&
     lngcolors(2&, 3&) = lngblue2&
     lngcolors(3&, 1&) = lngred3&
     lngcolors(3&, 2&) = lnggreen3&
     lngcolors(3&, 3&) = lngblue3&
     lngcolors(4&, 1&) = lngred4&
     lngcolors(4&, 2&) = lnggreen4&
     lngcolors(4&, 3&) = lngblue4&
    ReDim lngcolors2(1& To 4&, 1& To 3&) As Double
     obj100th = theobject.ScaleHeight / 100&
     lngswidth& = theobject.ScaleWidth
    theobject.Cls
    For index& = 1& To 4&
        For index2& = 1& To 3&
             lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
     theobject.BackColor = RGB(lngcolors2(4&, 1&), lngcolors2(4&, 2&), lngcolors2(4&, 3&))
    For index& = 1& To (4& - 1&)
         lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (4& - 1&))
         lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (4& - 1&))
         lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (4& - 1&))
        For index2& = 1& To (100& / (4& - 1&))
            theobject.Line (0&, lnghold)-(lngswidth&, lnghold + obj100th), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
             lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
             lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
             lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
             lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub

Public Sub fadepicbox(picbox As Object, color1 As Long, color2 As Long)
    'this goes in the _paint event of
    'the picture box
    Dim lngcon As Long, longcon As Long, lnghlfwidth As Long, lngcolorval1 As Long
    Dim lngcolorval2 As Long, lngcolorval3 As Long, lngrgb1 As Long, lngrgb2 As Long
    Dim lngrgb3 As Long, lngyval As Long, strcolor1 As String, strcolor2 As String
    Dim strred1 As String, strgreen1 As String, strblue1 As String, strred2 As String
    Dim strgreen2 As String, strblue2 As String, lngred1 As Long, lnggreen1 As Long
    Dim lngblue1 As Long, lngred2 As Long, lnggreen2 As Long, lngblue2 As Long
     picbox.AutoRedraw = True
     picbox.DrawStyle = 6&
     picbox.DrawMode = 13&
     picbox.DrawWidth = 2&
     lngcon& = 0&
     lnghlfwidth& = picbox.Width / 2&
     strcolor1$ = gethexfromrgb(color1&)
     strcolor2$ = gethexfromrgb(color2&)
     strred1$ = "&h" & Right$(strcolor1$, 2&)
     strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
     strblue1$ = "&h" & Left$(strcolor1$, 2&)
     strred2$ = "&h" & Right$(strcolor2$, 2&)
     strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
     strblue2$ = "&h" & Left$(strcolor2$, 2&)
     lngred1& = Val(strred1$)
     lnggreen1& = Val(strgreen1$)
     lngblue1& = Val(strblue1$)
     lngred2& = Val(strred2$)
     lnggreen2& = Val(strgreen2$)
     lngblue2& = Val(strblue2$)
    Do: DoEvents
        On Error Resume Next
         lngcolorval1& = lngred2& - lngred1&
         lngcolorval2& = lnggreen2& - lnggreen1&
         lngcolorval3& = lngblue2& - lngblue1&
         lngrgb1& = (lngcolorval1& / lnghlfwidth& * lngcon&) + lngred1&
         lngrgb2& = (lngcolorval2& / lnghlfwidth& * lngcon&) + lnggreen1&
         lngrgb3& = (lngcolorval3& / lnghlfwidth& * lngcon&) + lngblue1&
        picbox.Line (lngyval&, 0&)-(lngyval& + 2&, picbox.Height), RGB(lngrgb1&, lngrgb2&, lngrgb3&), BF
         lngyval& = lngyval& + 10&
         lngcon& = lngcon& + 5&
    Loop Until lngcon& > lnghlfwidth&
End Sub

Public Function fadetext2bybars(Red1 As Long, Green1 As Long, Blue1 As Long, red2 As Long, green2 As Long, blue2 As Long, strText As String, wavyonoff As Boolean) As String
    'IDEA from monkefade3
    'this will fade text by 2 colors
    'using the value of scroll bars
    Dim lnglen As Long, strcurstring As String, strlastchar As String
    Dim lngrgbcolor As Long, lnghexcolor As String, wherewavy As Long
    Dim wavystring As String, stroutput As String, index As Long
     lnglen& = Len(strText$)
    For index& = 1& To lnglen&
         strcurstring$ = Left$(strText$, index&)
         strlastchar$ = Right$(strcurstring$, 1&)
         lngrgbcolor& = RGB(((blue2& - Blue1&) / lnglen& * index&) + Blue1&, ((green2& - Green1&) / lnglen& * index&) + Green1&, ((red2& - Red1&) / lnglen& * index&) + Red1&)
         lnghexcolor$ = gethexfromrgb(lngrgbcolor&)
        If wavyonoff = True Then
             wherewavy& = wherewavy& + 1&
            Select Case wherewavy
                Case Is > 4&
                     wherewavy& = 1&
                Case 1&
                     wavystring$ = "<sup>"
                Case 2&
                     wavystring$ = "</sup>"
                Case 3&
                     wavystring$ = "<sub>"
                Case 4&
                     wavystring$ = "</sub>"
                Case Else
                    Exit Function
            End Select
        Else
             wavystring$ = ""
        End If
         stroutput$ = stroutput$ & "<Font Color=#" & lnghexcolor$ & ">" & wavystring$ & strlastchar$
    Next index&
     fadetext2bybars$ = stroutput$
End Function
Public Function fadetext2byrgb(color1 As Long, color2 As Long, strText As String, wavyonoff As Boolean) As String
    'IDEA from monkefade3
    'this will fade text, using 2 colors
    'you must supply the sub with the two colors
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long
    If strText$ = "" Then
        Exit Function
    Else
         strcolor1$ = gethexfromrgb(color1&)
         strcolor2$ = gethexfromrgb(color2&)
         strred1$ = "&H" & Right$(strcolor1$, 2&)
         strgreen1$ = "&H" & Mid$(strcolor1$, 3&, 2&)
         strblue1$ = "&H" & Left$(strcolor1$, 2&)
         strred2$ = "&H" & Right$(strcolor2$, 2&)
         strgreen2$ = "&H" & Mid$(strcolor2$, 3&, 2&)
         strblue2$ = "&H" & Left$(strcolor2$, 2&)
         lngred1& = Val(strred1$)
         lnggreen1& = Val(strgreen1$)
         lngblue1& = Val(strblue1$)
         lngred2& = Val(strred2$)
         lnggreen2& = Val(strgreen2$)
         lngblue2& = Val(strblue2$)
         fadetext2byrgb$ = fadetext2bybars(lngred1&, lnggreen1&, lngblue1&, lngred2&, lnggreen2&, lngblue2&, strText$, wavyonoff)
    End If
End Function

Public Function fadetext3bybars(Red1 As Long, Green1 As Long, Blue1 As Long, red2 As Long, green2 As Long, blue2 As Long, red3 As Long, green3 As Long, blue3 As Long, strText As String, wavyonoff As Boolean)
    'IDEA from monkefade3
    'this will fade text by 3 colors
    'using the value of scroll bars
    Dim lnglen As Long, strfirsthalf As String, strsecondhalf As String, index As Long
    Dim strcurstring As String, strlastchar As String, lngrgbcolor As Long, lnghexcolor As String
    Dim stroutput2 As String, stroutput1 As String, wherewavy As Long, wavystring As String
     lnglen& = Len(strText$)
     strfirsthalf$ = Left$(strText$, Int(lnglen&) / 2&)
     strsecondhalf$ = Right$(strText$, lnglen& - Int(lnglen&) / 2&)
    For index& = 1& To Len(strfirsthalf$)
         strcurstring$ = Left$(strfirsthalf$, index&)
         strlastchar$ = Right$(strcurstring$, 1&)
         lngrgbcolor& = RGB(((blue2& - Blue1&) / Len(strfirsthalf$) * index&) + Blue1&, ((green2& - Green1&) / Len(strfirsthalf$) * index&) + Green1&, ((red2& - Red1&) / Len(strfirsthalf$) * index&) + Red1&)
         lnghexcolor$ = gethexfromrgb(lngrgbcolor&)
        If wavyonoff = True Then
             wherewavy& = wherewavy& + 1&
            Select Case wherewavy
                Case Is > 4&
                     wherewavy& = 1&
                Case 1&
                     wavystring$ = "<sup>"
                Case 2&
                     wavystring$ = "</sup>"
                Case 3&
                     wavystring$ = "<sub>"
                Case 4&
                     wavystring$ = "</sub>"
                Case Else
                    Exit Function
            End Select
        Else
             wavystring$ = ""
        End If
         stroutput1$ = stroutput1$ & "<Font Color=#" & lnghexcolor$ & ">" & wavystring$ & strlastchar$
    Next index&
    For index& = 1 To Len(strsecondhalf$)
         strcurstring$ = Left$(strsecondhalf$, index&)
         strlastchar$ = Right$(strcurstring$, 1&)
         lngrgbcolor& = RGB(((blue2& - Blue1&) / Len(strsecondhalf$) * index&) + Blue1&, ((green2& - Green1&) / Len(strsecondhalf$) * index&) + Green1&, ((red2& - Red1&) / Len(strsecondhalf$) * index&) + Red1&)
         lnghexcolor$ = gethexfromrgb(lngrgbcolor&)
        If wavyonoff = True Then
             wherewavy& = wherewavy& + 1&
            Select Case wherewavy
                Case Is > 4&
                     wherewavy& = 1&
                Case 1&
                     wavystring$ = "<sup>"
                Case 2&
                     wavystring$ = "</sup>"
                Case 3&
                     wavystring$ = "<sub>"
                Case 4&
                     wavystring$ = "</sub>"
                Case Else
                    Exit Function
            End Select
        Else
             wavystring$ = ""
        End If
         stroutput2$ = "<Font Color=#" & lnghexcolor$ & ">" & wavystring$ & strlastchar$ & stroutput2$
    Next index&
     fadetext3bybars = stroutput1$ & stroutput2$
End Function

Public Function fadetext3byrgb(color1 As Long, color2 As Long, color3 As Long, strText As String, wavyonoff As Boolean) As String
    'this will fade text, using 3 colors
    'you must supply the sub with the three colors
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    If strText$ = "" Then
        Exit Function
    Else
         strcolor1$ = gethexfromrgb(color1&)
         strcolor2$ = gethexfromrgb(color2&)
         strcolor3$ = gethexfromrgb(color3&)
         strred1$ = "&H" & Right$(strcolor1$, 2&)
         strgreen1$ = "&H" & Mid$(strcolor1$, 3&, 2&)
         strblue1$ = "&H" & Left$(strcolor1$, 2&)
         strred2$ = "&H" & Right$(strcolor2$, 2&)
         strgreen2$ = "&H" & Mid$(strcolor2$, 3&, 2&)
         strblue2$ = "&H" & Left$(strcolor2$, 2&)
         strred3$ = "&H" & Right$(strcolor3$, 2&)
         strgreen3$ = "&H" & Mid$(strcolor3$, 3&, 2&)
         strblue3$ = "&H" & Left$(strcolor3$, 2&)
         lngred1& = Val(strred1$)
         lnggreen1& = Val(strgreen1$)
         lngblue1& = Val(strblue1$)
         lngred2& = Val(strred2$)
         lnggreen2& = Val(strgreen2$)
         lngblue2& = Val(strblue2$)
         fadetext3byrgb$ = fadetext3bybars(lngred1&, lnggreen1&, lngblue1&, lngred2&, lnggreen2&, lngblue2&, lngred3&, lnggreen3&, lngblue3&, strText$, wavyonoff)
    End If
End Function

Public Sub deeoldmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, lngbutton As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
             lngbutton& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub

Public Sub ctrlaltdel(enabled As Boolean)
    Dim lnggogo As Long, pOld As Boolean
     lnggogo& = SystemParametersInfo(SPI_SCREENSAVERRUNNING, enabled, pOld, 0&)
End Sub

Public Sub decompileprotect(strexepathandname As String)
    Dim strFile As String
    On Error Resume Next
    If Not InStr(strexepathandname$, "\") Then
        MsgBox "executable file not found", vbOKOnly, "izekial32.bas": Exit Sub
    Else
         strFile$ = FreeFile
        Open strexepathandname$ For Binary As #strFile$
            Seek #strFile$, 25&

        Close #strFile$
        If Err Then
            MsgBox "not a visual basic made file!", vbOKOnly, "error in file": Exit Sub
        Else
            MsgBox "the file below has been decompile protected" & vbCrLf$ & strexepathandname$, vbOKOnly, "izekial32.bas"
        End If
    End If
End Sub
Public Sub deeallflashmail()
    Dim index As Long
    For index& = 0& To getflashmailcount& - 1&
        Call deeflashmailitem(index&)
    Next index&
End Sub

Public Sub deeallflashmailwithdogbar(progbar As Control)
    Dim index As Long
     progbar.Max = getflashmailcount& - 1&
    For index& = 0& To getflashmailcount& - 1&
         progbar.Value = index& - 1&
        Call deeflashmailitem(index&)
    Next index&
End Sub

Public Sub deeallnewmail()
    Dim index As Long
    For index& = 0& To getnewmailcount&
        Call deenewmailitem(index&)
    Next index&
End Sub

Public Sub deeallnewmailwithdogbar(progbar As Control)
    Dim index As Long
     progbar.Max = getnewmailcount& - 1&
    For index& = 0& To getnewmailcount& - 1&
         progbar.Value = index& - 1&
        Call deenewmailitem(index&)
    Next index&
End Sub

Public Sub deealloldmail()
    Dim index As Long
    For index& = 0& To getoldmailcount& - 1&
        Call deeoldmailitem(index&)
    Next index&
End Sub

Public Sub deeallsentmail()
    Dim index As Long
    For index& = 0& To getsentmailcount& - 1&
        Call deesentmailitem(index&)
    Next index&
End Sub

Public Sub deeallsentmailwithdogbar(progbar As Control)
    Dim index As Long
     progbar.Max = getsentmailcount& - 1&
    For index& = 0& To getsentmailcount& - 1&
         progbar.Value = index& - 1&
        Call deesentmailitem(index&)
    Next index&
End Sub

Public Sub deeflashmailitem(index As Long)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lnglist As Long
    Dim lngdelbut As Long, lngokwin As Long, lngokbut As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
     lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
     lnglist& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
    If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& < index& Then
        Exit Sub
    Else
         lngdelbut& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
         lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
         lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
         lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngdelbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngdelbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngokwin& = FindWindow("#32770", vbNullString)
             lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "&Yes")
        Loop Until lngokbut& <> 0&
        Do: DoEvents
             lngokwin& = FindWindow("#32770", vbNullString)
             lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "&Yes")
            Call SendMessageLong(lngokbut&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessageLong(lngokbut&, WM_KEYUP, VK_SPACE, 0&)
        Loop Until lngokbut& = 0&
    End If
End Sub








Sub CenterForm(Frm As Form)
'umm yeah   i think you know what this does..
Frm.Top = (Screen.Height * 0.3) / 2 - (Frm.Height / 2)
Frm.Left = (Screen.Width / 1) - (Frm.Width / 1)
End Sub


Sub chr_convert()
'Dim strin As String
'strin = convert.Text
'If Len(strin) = 0 Then response = MsgBox("Error!" & Chr(13) & "You must enter in a string to convert!", 64, "Error - No String"): Exit Sub
'For i = 1 To Len(strin)
'cahr$ = Mid$(strin, i, 1)
'mmmcrap = Asc(cahr$)
'If i = 1 Then harlem = "Chr(" & mmmcrap & ")" Else harlem = harlem & " & Chr(" & mmmcrap & ")"
'Next i
'txtCode = harlem
'hidden = harlem
End Sub


Sub Clip_Copyto(what$)
'set text to the clipboard
Clipboard.SetText what$
End Sub

Sub clip_cuttext(Text1 As TextBox)
Clipboard.SetText Text1.Text
Text1.Text = ""
Text1.Text = ""
MsgBox "The ascii codes have been cut to the clipboard.", 64, "Codes cut."

End Sub

Sub ControlPanels(FileName As String)
'Sub declared by Tech
Dim rtn As Double
On Error Resume Next
rtn = Shell(FileName, 5)
End Sub

Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub


Public Sub deesentmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, lngbutton As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
             lngbutton& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
             lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub


Public Function getflashmailcount() As Long
    Dim lngaol As Long, lngmdi As Long, lngmailwin As Long, lnglist As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
     lngmailwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
     lnglist& = FindWindowEx(lngmailwin&, 0&, "_AOL_Tree", vbNullString)
     getflashmailcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function getflashmailsubject(index As Long) As String
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lnglist As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
     lnglist& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
    If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) < index& - 1& Then
        Exit Function
    Else
        If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
             lngtextlen& = SendMessage(lnglist&, LB_GETTEXTLEN, index&, 0&)
             strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lnglist&, LB_GETtext, index&, strmailtext$)
             lngtabkey& = InStr(strmailtext$, Chr$(9&))
             lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
             strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
             strmailtext$ = ReplaceString(strmailtext$, Chr$(0&), "")
             getflashmailsubject$ = strmailtext$
        End If
    End If
End Function

Public Function getfontcount() As Long
     getfontcount& = Screen.FontCount
End Function

Public Function gethexfromrgb(rgbvalue As Long) As String
    Dim hexstate As String, hexlen As Long
     hexstate$ = Hex(rgbvalue&)
     hexlen& = Len(hexstate$)
    Select Case hexlen&
        Case 1&
             gethexfromrgb$ = "00000" & hexstate$
            Exit Function
        Case 2&
             gethexfromrgb$ = "0000" & hexstate$
            Exit Function
        Case 3&
             gethexfromrgb$ = "000" & hexstate$
            Exit Function
        Case 4&
             gethexfromrgb$ = "00" & hexstate$
            Exit Function
        Case 5&
             gethexfromrgb$ = "0" & hexstate$
            Exit Function
        Case 6&
             gethexfromrgb$ = "" & hexstate$
            Exit Function
        Case Else
            Exit Function
    End Select
End Function


Sub CPanel_Commands()
'Sub declared by Tech

'Add New Hardware
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1")

'Add/Remove Programs
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1")

'Time/Date Controls
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")

'Display Settings
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0")

'Internet Properties
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0")

'JoyStick Controls
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL joy.cpl")

'KeyBoard Settings
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1")

'Modem Preferences
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL modem.cpl")

'Mouse Settings
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0")

'MultiMedia
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0")

'Network
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl")

'Passwords
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL password.cpl")

'Regional Settings
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0")

'Sounds
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1")

'System Properties
Call ControlPanels("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0")

End Sub


Public Function RemSpaces(ByVal bvnc As String)
'This function will remove all spaces from a string
Dim x
Dim Total$
'nospace$ = RemSpaces("Joe V 456")
'in this example nospaces$ = "JoeV456"
For x = 1 To Len(bvnc)
    If Mid$(bvnc, x, 1) <> Space(1) Then Total$ = Total$ + Mid$(bvnc, x, 1)
Next x
RemSpaces = Total$
End Function



Public Sub aolgoforward()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Function ReplaceString(ByVal Strt As String, ByVal ReplaceMe As String, ByVal ReplaceWith As String) As String
'This function will replace a string in a string
'with another string =)
Dim Start$
Dim x%
Dim pos%

'after$ = ReplaceString("its a boy", "boy", "girl")
'in this example after$ = "its a girl"
Start$ = Strt
Do While InStr(Start$, ReplaceMe$) <> 0
    x% = DoEvents()
    pos% = InStr(Start$, ReplaceMe$)
    Start$ = Left(Start$, pos% - 1) & ReplaceWith$ & Right(Start$, Len(Start$) - pos% - Len(ReplaceMe$) + 1)
    Loop
ReplaceString$ = Start$
End Function

Sub SaveINI(Title$, TheKey$, TheValue$, ThePath$)
 Dim h%
    h% = WritePrivateProfileString(Title$, TheKey$, TheValue$, ThePath$)
End Sub

Sub scan_BO_Registrycheck1()
Dim a!
a! = Shell("c:\windows\scanregw.exe")
End Sub


Public Function getnewmailcount() As Long
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
         getnewmailcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
    End If
End Function
Public Function getnewmailsubject(index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
             lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
             strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
             lngtabkey& = InStr(strmailtext$, Chr$(9&))
             lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
             strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
             strmailtext$ = ReplaceString(strmailtext$, Chr$(0&), "")
             getnewmailsubject$ = strmailtext$
        End If
    End If
End Function



Public Function GetHostByAddress(ByVal addr As Long) As String
    Dim phe&, ret&
    Dim heDestHost As HostEnt
    Dim HostName$
    phe = gethostbyaddr(addr, 4, PF_INET)
    If phe Then
        MemCopy heDestHost, ByVal phe, hostent_size
        HostName = String(256, 0)
        MemCopy ByVal HostName, ByVal heDestHost.h_name, 256
        GetHostByAddress = Left(HostName, InStr(HostName, Chr(0)) - 1)
    Else
        GetHostByAddress = WSA_NoName
    End If
End Function


'returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal HostName$) As Long
    'Return IP address as a long, in network byte order
    Dim phe&
    Dim heDestHost As HostEnt
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(HostName$)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(HostName$)
        If phe <> 0 Then
            MemCopy heDestHost, ByVal phe, hostent_size
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
End Function

'returns your local machines name
Public Function GetLocalHostName() As String
    Dim sName$
    sName = String(256, 0)
    If gethostname(sName, 256) Then
        sName = WSA_NoName
    Else
        If InStr(sName, Chr(0)) Then
            sName = Left(sName, InStr(sName, Chr(0)) - 1)
        End If
    End If
    GetLocalHostName = sName
End Function

Public Function getlistcount(lnglist As Long) As Long
     getlistcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function getlistitemindex(TheList As Long, Item As String) As Long
    On Error Resume Next
    Dim rList As Long, sThread As Long, mThread As Long, index As Long
    Dim ScreenName As String, itmHold As Long, psnHold As Long
    Dim rBytes As Long, cProcess As Long
     rList& = TheList&
     sThread& = GetWindowThreadProcessId(rList&, cProcess&)
     mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0& To SendMessage(rList&, LB_GETCOUNT, 0&, 0&) - 1&
             ScreenName$ = String$(4&, vbNullChar)
             itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
             itmHold& = itmHold& + 24&
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4&, rBytes&)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4&)
             psnHold& = psnHold& + 6&
             ScreenName$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
             ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1&)
            If ScreenName$ = Item$ Then
                 getlistitemindex& = index&
                Call CloseHandle(mThread&)
                Exit Function
            End If
        Next index&
         getlistitemindex& = -1&
        Call CloseHandle(mThread&)
    End If
End Function


Public Function getoldmailcount() As Long
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
         getoldmailcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
    End If
End Function

Public Function getoldmailsubject(index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
             lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
             strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
             lngtabkey& = InStr(strmailtext$, Chr$(9&))
             lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
             strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
             strmailtext$ = ReplaceString(strmailtext$, Chr$(0&), "")
             getoldmailsubject$ = strmailtext$
        End If
    End If
End Function

Public Function getprofile(ByVal strperson As String) As String
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim mousecur As POINTAPI, lngprofwin As Long, lngedit As Long, lngbutton As Long
    Dim lngview As Long, lngnoprofwin As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(mousecur.x, mousecur.Y)
    Do: DoEvents
         lngprofwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Get a Member's Profile")
         lngedit& = FindWindowEx(lngprofwin&, 0&, "_AOL_Edit", vbNullString)
         lngicon& = FindWindowEx(lngprofwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngprofwin& <> 0& And lngedit& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strperson$)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngprofwin&, WM_CLOSE, 0&, 0&)
    timeout 3
    Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(RICHCNTL, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(RICHCNTL, WM_gettext, TL + 1, TheText)
TheText = Left(TheText, TL)
getprofile = TheText

aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolstatic = FindWindowEx(aolchild, 0&, "_aol_static", vbNullString)
aolstatic = FindWindowEx(aolchild, aolstatic, "_aol_static", vbNullString)
            
            If aolstatic <> 0 And getprofile = "" Then GoTo done
                If getprofile = "" Then
                Call waitforok
                Exit Function
            End If
done:
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Call SendMessageLong(aolchild, WM_CLOSE, 0&, 0&)

End Function

Public Function getsentmailcount() As Long
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
         getsentmailcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
    End If
End Function

Public Function getsentmailsubject(index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
             lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
             strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETtext, index&, strmailtext$)
             lngtabkey& = InStr(strmailtext$, Chr$(9&))
             lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
             strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
             strmailtext$ = ReplaceString(strmailtext$, Chr$(0&), "")
             getsentmailsubject$ = strmailtext$
        End If
    End If
End Function

Public Function getstringlinecount(strstring As String) As Long
    Dim enterpos As Long
    If Len(strstring$) = 0& Then
         getstringlinecount& = 0&
    Else
         enterpos& = InStr(strstring$, Chr$(13&))
        If enterpos& <> 0& Then
             getstringlinecount& = 1&
            Do While enterpos& <> 0&
                 enterpos& = InStr(enterpos& + 1&, strstring$, Chr$(13&))
                If enterpos& <> 0& Then
                     getstringlinecount& = getstringlinecount& + 1&
                End If
            Loop: DoEvents
        End If
         getstringlinecount& = getstringlinecount& + 1&
    End If
End Function
Public Function getstringlineindex(strstring As String, strlinetext As String) As Long
    Dim index As Long
    For index& = 0& To getstringlinecount(strstring$)
        If getlinefromstring(strstring$, index&) = strlinetext$ Then
             getstringlineindex& = index&
            Exit Function
        End If
    Next index&
End Function

Public Function GetText(lngwindow As Long) As String
    Dim strBuffer As String, lngtextlen As Long
     lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
     strBuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_gettext, lngtextlen& + 1&, strBuffer$)
     GetText$ = strBuffer$
End Function

Public Function gettreecount(lngtree As Long) As Long
     gettreecount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function gettreeitemindex(lngtree As Long, stritem As String) As Long
    Dim index As Long, lngtextlen As Long, strstring As String
    For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
         lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
         strstring$ = String$(lngtextlen& + 1&, 0&)
        Call SendMessageByString(lngtree&, LB_GETtext, index&, strstring$)
        If strstring$ = stritem$ Then
             gettreeitemindex& = index&
            Exit Function
        End If
    Next index&
End Function

Public Function gettreeitemtext(lngtree As Long, index As Long) As String
    Dim lngtextlen As Long
     lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
     gettreeitemtext$ = String$(lngtextlen& + 1&, 0&)
    Call SendMessageByString(lngtree&, LB_GETtext, index&, gettreeitemtext$)
End Function

Public Function GetUser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    GetUser$ = ""
End Function



Public Sub ghost(ghost As Boolean)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long
    Dim lngsetupbut As Long, lngsetupwin As Long, lngpprefbut1 As Long
    Dim lngpprefbut2 As Long, lngpprefbut3 As Long, lngpprefbut4 As Long
    Dim lngpprefbut As Long, lngpprefswin As Long, lngblock1 As Long, buddyopen As Boolean
    Dim lngblock2 As Long, lngblock3 As Long, lngblock4 As Long, lngblock5 As Long
    Dim lngblock6 As Long, lngblock As Long, lngsavebut1 As Long, lngsavebut2 As Long
    Dim lngsavebut3 As Long, lngsavebut As Long, lngokwin As Long, lngokbut As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngbuddywin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
    If lngbuddywin& <> 0& Then
         buddyopen = True
    Else
        Call Keyword("bv")
         buddyopen = False
    End If
    Do: DoEvents
         lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_AOL_Icon", vbNullString)
         lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_AOL_Icon", vbNullString)
         lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_AOL_Icon", vbNullString)
        Call PostMessage(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
         lngsetupwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
         lngpprefbut1& = FindWindowEx(lngsetupwin&, 0&, "_AOL_Icon", vbNullString)
         lngpprefbut2& = FindWindowEx(lngsetupwin&, lngpprefbut1&, "_AOL_Icon", vbNullString)
         lngpprefbut3& = FindWindowEx(lngsetupwin&, lngpprefbut2&, "_AOL_Icon", vbNullString)
         lngpprefbut4& = FindWindowEx(lngsetupwin&, lngpprefbut3&, "_AOL_Icon", vbNullString)
         lngpprefbut& = FindWindowEx(lngsetupwin&, lngpprefbut4&, "_AOL_Icon", vbNullString)
    Loop Until lngsetupwin& <> 0& And lngpprefbut1& <> 0& And lngpprefbut2& <> 0& And lngpprefbut3& <> 0& And lngpprefbut4& <> 0& And _
    lngpprefbut1& <> lngpprefbut2& And lngpprefbut2& <> lngpprefbut3& And lngpprefbut3& <> lngpprefbut4& And lngpprefbut4& <> lngpprefbut&: Pause 0.2
    Do: DoEvents
        Call PostMessage(lngpprefbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngpprefbut&, WM_LBUTTONUP, 0&, 0&)
         lngpprefswin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Privacy Preferences")
         lngblock1& = FindWindowEx(lngpprefswin&, 0&, "_AOL_Checkbox", vbNullString)
         lngblock2& = FindWindowEx(lngpprefswin&, lngblock1&, "_AOL_Checkbox", vbNullString)
         lngblock3& = FindWindowEx(lngpprefswin&, lngblock2&, "_AOL_Checkbox", vbNullString)
         lngblock4& = FindWindowEx(lngpprefswin&, lngblock3&, "_AOL_Checkbox", vbNullString)
         lngblock5& = FindWindowEx(lngpprefswin&, lngblock4&, "_AOL_Checkbox", vbNullString)
         lngblock6& = FindWindowEx(lngpprefswin&, lngblock5&, "_AOL_Checkbox", vbNullString)
         lngblock& = FindWindowEx(lngpprefswin&, lngblock6&, "_AOL_Checkbox", vbNullString)
         lngsavebut1& = FindWindowEx(lngpprefswin&, 0&, "_AOL_Icon", vbNullString)
         lngsavebut2& = FindWindowEx(lngpprefswin&, lngsavebut1&, "_AOL_Icon", vbNullString)
         lngsavebut3& = FindWindowEx(lngpprefswin&, lngsavebut2&, "_AOL_Icon", vbNullString)
         lngsavebut& = FindWindowEx(lngpprefswin&, lngsavebut3&, "_AOL_Icon", vbNullString)
    Loop Until lngblock1& <> 0& And lngblock2& <> 0& And lngblock3& <> 0& And lngblock4& <> 0& And lngblock5& <> 0& And lngblock6& <> 0& And lngblock& <> 0& And lngsavebut& <> 0& And lngsavebut1& <> 0& And lngsavebut2& <> 0& And lngsavebut3& <> 0& And _
    lngblock1& <> lngblock2& And lngblock2& <> lngblock3& And lngblock3& <> lngblock4& And lngblock4& <> lngblock5 And lngblock5& <> lngblock6& And lngblock6& <> lngblock& And lngsavebut1& <> lngsavebut2& And lngsavebut2& <> lngsavebut3& And lngsavebut3& <> lngsavebut&
    If ghost = True Then
        Call PostMessage(lngblock5&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngblock5&, WM_LBUTTONUP, 0&, 0&)
    ElseIf ghost = False Then
        Call PostMessage(lngblock3&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngblock3&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call PostMessage(lngblock&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngblock&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Call PostMessage(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
         lngokwin& = FindWindow("#32770", "America Online")
         lngokbut& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Pause 0.2
    Loop Until lngokwin& <> 0& And lngokbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
         lngokwin& = FindWindow("#32770", "America Online")
    Loop Until lngokwin& = 0&: Pause 0.2
    Call PostMessage(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub IMIgnore(strwho, ignore As Boolean)
    Dim strstate As String
    If ignore = True Then
         strstate$ = "off"
    Else
         strstate$ = "on"
    End If
    Call SendIM("$im_" & strstate$ & " " & strwho, " ")
End Sub

Public Sub InstantMessage(Person As String, Message As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, ok As Long, button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        ok& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until ok& <> 0& Or im& = 0&
    If ok& <> 0& Then
        button& = FindWindowEx(ok&, 0&, "Button", vbNullString)
        Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Function isalive(strName As String) As Boolean
    Dim lngaol As Long, lngmdi As Long, lngerrorwin As Long
    Dim lngerrorview As Long, strerrorstring As String
    Dim lngmailwin As Long, lngnowin As Long, lngbutton As Long
    Call SendMail("*, " & strName$, "you alive?", " ")
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
         lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
         lngerrorview& = FindWindowEx(lngerrorwin&, 0&, "_AOL_View", vbNullString)
         strerrorstring$ = GetText(lngerrorview&)
    Loop Until lngerrorwin& <> 0& And lngerrorview& <> 0& And strerrorstring$ <> ""
    If InStr(LCase$(ReplaceString(strerrorstring$, " ", "")), LCase$(ReplaceString(strName$, " ", ""))) > 0& Then
         isalive = False
    Else
         isalive = True
    End If
     lngmailwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(lngmailwin&, WM_CLOSE, 0&, 0&)
    Call PostMessage(lngerrorwin&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do: DoEvents
         lngnowin& = FindWindow("#32770", "America Online")
         lngbutton& = FindWindowEx(lngnowin&, 0&, "Button", "&No")
    Loop Until lngnowin& <> 0& And lngbutton& <> 0
    Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
End Function



Public Function isavailable(strName As String) As Boolean
    Dim lngaol As Long, lngmdi As Long, lngimwin As Long, lngrich As Long, lngokwin As Long
    Dim lngicon As Long, lngbut As Long, lngstatic As Long, strmsg As String, lngedit As Long
     lngaol& = FindWindow("AOL Frame25", "America  Online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & strName$)
    Do: DoEvents
         lngimwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Send Instant Message")
         lngrich& = FindWindowEx(lngimwin&, 0&, "RICHCNTL", vbNullString)
         lngedit& = FindWindowEx(lngimwin&, 0&, "_AOL_Edit", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, 0&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngimwin& <> 0& And lngicon& <> 0& And lngrich& <> 0& And lngedit& <> 0&
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngokwin& = FindWindow("#32770", "America Online")
         lngbut& = FindWindowEx(lngokwin&, 0&, "Button", vbNullString)
         lngstatic& = FindWindowEx(lngokwin&, 0&, "Static", vbNullString)
         lngstatic& = FindWindowEx(lngokwin&, lngstatic&, "Static", vbNullString)
    Loop Until lngokwin& <> 0& And lngbut& <> 0& And lngstatic& <> 0&
     strmsg$ = GetText(lngstatic&)
    Do: DoEvents
        Call PostMessage(lngbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngbut&, WM_LBUTTONUP, 0&, 0&)
         lngokwin& = FindWindow("#32770", "America Online")
         lngbut& = FindWindowEx(lngokwin&, 0&, "Button", vbNullString)
    Loop Until lngokwin& = 0& And lngbut& = 0&
    Call SendMessage(lngimwin&, WM_CLOSE, 0&, 0&)
    If LCase$(strmsg$) = LCase$(strName$) & " is online and able to receive instant messages." Then
         isavailable = True
    Else
         isavailable = False
    End If
End Function

Public Function ismaster() As Boolean
    Dim lngaol As Long, lngmdi As Long, lngprntwindow As Long
    Dim lngprnticon As Long, lngmodal As Long, lngmodalstatic As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Call Keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do Until lngprntwindow& <> 0& And lngprnticon& <> 0&
         lngprntwindow& = FindWindowEx(lngmdi&, 0&, "AOL Child", " Parental Controls")
         lngprnticon& = FindWindowEx(lngprntwindow&, 0&, "_AOL_Icon", vbNullString)
    Loop: DoEvents
    Do: DoEvents
        Call PostMessage(lngprnticon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngprnticon&, WM_LBUTTONUP, 0&, 0&)
         lngmodal& = FindWindow("_AOL_Modal", vbNullString)
         lngmodalstatic& = FindWindowEx(lngmodal&, 0&, "_AOL_Static", vbNullString)
    Loop Until lngmodal& <> 0 And lngmodalstatic& <> 0& And GetText(lngmodalstatic&) <> ""
    If ReplaceString(ReplaceString(GetText(lngmodalstatic&), Chr$(10&), ""), Chr$(13&), "") = "Set Parental Controls" Then
         ismaster = True
    Else
         ismaster = False
    End If
    Call SendMessageLong(lngmodal&, WM_CLOSE, 0&, 0&): DoEvents
    Call SendMessageLong(lngprntwindow&, WM_CLOSE, 0&, 0&): DoEvents
End Function

Public Function isonline() As Boolean
    Dim lngaol As Long, lngmdi As Long, lngtoolbarwin As Long, lngtoolbarico As Long
    Dim lngcomposeico As Long
     lngaol& = FindWindow("aol frame25", "america  online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbarwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
     lngtoolbarico& = FindWindowEx(lngtoolbarwin&, 0&, "_aol_toolbar", vbNullString)
     lngcomposeico& = FindWindowEx(lngtoolbarico&, 0&, "_aol_icon", vbNullString)
     lngcomposeico& = FindWindowEx(lngtoolbarico&, lngcomposeico&, "_aol_icon", vbNullString)
    If lngcomposeico& <> 0& Then
         isonline = True
    Else
         isonline = False
    End If
End Function

Public Sub Keyword(KW As String)
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Function isnumber(blah)
Dim zoink
Dim a
Dim b
Dim thechr
zoink = 1
For a = 1 To Len(blah)
thechr = LCase$(Right$(Left$(blah, a), 1))
   b = 0
   If thechr = "0" Then b = 1
   If thechr = "1" Then b = 1
   If thechr = "2" Then b = 1
   If thechr = "3" Then b = 1
   If thechr = "4" Then b = 1
   If thechr = "5" Then b = 1
   If thechr = "6" Then b = 1
   If thechr = "7" Then b = 1
   If thechr = "8" Then b = 1
   If thechr = "9" Then b = 1
   If b = 0 Then zoink = 0
Next a
 If zoink = 0 Then
    isnumber = 0
 Else
    isnumber = 1
 End If
End Function

Public Sub killmailad(lngmailbox As Long)
    'ex: call killmailad(findnewmailbox&)
    Dim lngadv As Long
     lngadv& = FindWindowEx(lngmailbox&, 0&, "_AOL_Image", vbNullString)
    Call SendMessage(lngadv&, WM_CLOSE, 0&, 0&)
End Sub
Function isletter(a)
Dim b
Dim c
Dim d
Dim e
b = 1
a = LCase$(nospaces("" & a))
For c = 1 To Len(a)
d = LCase$(Right$(Left$(a, c), 1))
   e = 0
   If d = "a" Then e = 1
   If d = "b" Then e = 1
   If d = "c" Then e = 1
   If d = "d" Then e = 1
   If d = "e" Then e = 1
   If d = "f" Then e = 1
   If d = "g" Then e = 1
   If d = "h" Then e = 1
   If d = "i" Then e = 1
   If d = "j" Then e = 1
   If d = "k" Then e = 1
   If d = "l" Then e = 1
   If d = "m" Then e = 1
   If d = "n" Then e = 1
   If d = "o" Then e = 1
   If d = "p" Then e = 1
   If d = "q" Then e = 1
   If d = "r" Then e = 1
   If d = "s" Then e = 1
   If d = "t" Then e = 1
   If d = "u" Then e = 1
   If d = "v" Then e = 1
   If d = "w" Then e = 1
   If d = "x" Then e = 1
   If d = "y" Then e = 1
   If d = "z" Then e = 1
   If e = 0 Then b = 0
Next c
 If b = 0 Then
    isletter = 0
 Else
    isletter = 1
 End If

End Function

Sub KillWin(Windo)
Dim x
x = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Public Sub killwait()
    Dim lngmodal As Long, lngbut As Long
    Call runaolmenu(4&, 10&)
    Do: DoEvents
        lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        lngbut& = FindWindowEx(lngmodal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngmodal& <> 0& And lngbut& <> 0&
    Call PostMessage(lngbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngbut&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function listfindstring(List As Listbox, findstring As String) As Long
    Dim index As Long
    If List.ListCount = 0 Then Exit Function
    For index& = 0 To List.ListCount - 1
         List.ListIndex = index&
        If UCase(List.Text) = UCase(findstring$) Then
             listfindstring& = index&
            Exit Function
            If Err Then Exit Function
        End If
    Next index&
End Function
Public Sub LoadComboBox(strnameandpath As String, TheCombo As ComboBox)
    'this won't work with all saved combos
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    Open strnameandpath$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            TheCombo.AddItem strlinetext$
        Wend
    Close #1&
End Sub

Public Sub loadcomboboxwithcommondialog(commondialog32 As Control, TheCombo As TextBox)
    'this won't work with all saved combos
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    With commondialog32
        .DialogTitle = "load a file"
        .CancelError = True
        .FilterIndex = 0&
        .ShowOpen
    End With
    Open commondialog32.FileName For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            TheCombo.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Function nospaces(ByVal what As String)
'gets all the spaces out of text
Dim endstring
Dim i
Dim char
Dim TheStr

endstring = ""
For i = 1 To Len(what)
char = Mid(what, i, 1)
If char = " " Then char = ""
TheStr = TheStr & char
Next i
nospaces = TheStr

End Function

Public Sub Loadlistbox(strnameandpath As String, TheList As Listbox)
    'this won't work with all saved lists
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    Open strnameandpath$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            TheList.AddItem strlinetext$
        Wend
    Close #1&
End Sub



Public Sub loadmeonstartup()
    Call writeini("windows", "load", App.Path & "\" & App.EXEName, "c:\window\win.ini")
End Sub

Public Sub loadsnlist(TheList As Control, commondlg32 As Control)
    Dim strsnhold As String, strthecomma As String, strsn As String, index As Long
    With commondlg32
        .DialogTitle = "load a sn list"
        .CancelError = True
        .Filter = "text file (*.txt)|*.txt"
        .FilterIndex = 0&
        .ShowOpen
    End With
    Open commondlg32.FileName For Input As #1
         strsnhold$ = Input(LOF(1&), 1&)
    Close #1
    For index& = 1& To Len(strsnhold$)
         strthecomma$ = Mid$(strthecomma$, index&, 1)
        If strthecomma$ = "," Then
            Call TheList.AddItem(strsn$)
             strsn$ = ""
        Else
              strsn$ = strsn$ & strthecomma$
        End If
    Next index&
End Sub


Public Function locatemember(strName As String) As String
    Dim lngokwin As Long, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngchildcap As String, lnglocatewin As Long, lngstaticwin As Long
    Dim strstaticmsg As String, lngbutton As Long
    Call Keyword("aol://3548:" & strName$)
    Do: DoEvents
         lngokwin& = FindWindow("#32770", "America Online")
        If lngokwin& <> 0& Then
            Exit Do
        Else
             lngaol& = FindWindow("AOL Frame25", "America  Online")
             lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
             lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
             lngchildcap$ = GetCaption(lngchild&)
            If LCase$(lngchildcap$) = LCase$("locate " & strName$) Then
                 lnglocatewin& = lngchild&
            Else
                 lngaol& = FindWindow("AOL Frame25", "America  Online")
                 lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
                 lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
                Do: DoEvents
                     lngchildcap$ = GetCaption(lngchild&)
                    If LCase$(lngchildcap$) = LCase$("locate " & strName$) Then
                         lnglocatewin& = lngchild&
                        Exit Do
                    End If
                     lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
                     lngokwin& = FindWindow("#32770", "America Online")
                Loop Until lngchild& = 0& Or lngokwin& <> 0&
            End If
        End If
    Loop Until lnglocatewin& <> 0& Or lngokwin& <> 0&
    If lnglocatewin& <> 0& Then
         lngstaticwin& = FindWindowEx(lnglocatewin&, 0&, "_AOL_Static", vbNullString)
         strstaticmsg$ = GetText(lngstaticwin&)
        If LCase$(strstaticmsg$) = LCase$(strName$ & " is online, but not in a chat area.") Then
             locatemember$ = "online, but not in a chat room"
        ElseIf LCase$(strstaticmsg$) = LCase$(strName$ & " is online, but in a private room.") Then
             locatemember$ = "online, and in a private room"
        ElseIf LCase$(strstaticmsg$) Like LCase$(strName$ & " is in chat room *") Then
             locatemember$ = "online, and in " & Right$(strstaticmsg$, Len(strstaticmsg$) - Len(strName$) + 17&)
        End If
        Call SendMessage(lnglocatewin&, WM_CLOSE, 0&, 0&)
    ElseIf lngokwin& <> 0& Then
         lngbutton& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Do: DoEvents
            Call PostMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
             lngokwin& = FindWindow("#32770", "America Online")
             lngbutton& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Loop Until lngokwin& = 0& And lngbutton& = 0&
         locatemember$ = "not online."
    End If
End Function

Public Sub makeshortcut(strshortcutdir As String, strshortcutname As String, strshortcutpath As String)
    Dim strwinshortcutdir As String, strwinshortcutname As String, strwinshortcutexepath As String, lngretval As Long
     strwinshortcutdir$ = strshortcutdir$
     strwinshortcutname$ = strshortcutname$
     strwinshortcutexepath$ = strshortcutpath$
     lngretval& = fCreateShellLink("", strwinshortcutname$, strwinshortcutexepath$, "")
    Name "c:\windows\start menu\programs\" & strwinshortcutname$ & ".lnk" As strwinshortcutdir$ & "\" & strwinshortcutname$ & ".lnk"
End Sub

Public Sub MemberRoom(strroom As String)
    Call Keyword("aol://2719:61-2-" & strroom$)
End Sub

Public Sub midiplay(strmidi As String)
    Dim strFile As String
     strFile$ = Dir(strmidi$)
    If strFile$ = "" Then
        Exit Sub
    Else
        Call mciSendString("play " & strmidi$, 0&, 0, 0)
    End If
End Sub

Public Sub midistop(strmidi As String)
    Dim strFile As String
    strFile$ = Dir(strmidi$)
    If strFile$ = "" Then
        Exit Sub
    Else
        Call mciSendString("stop " & strmidi$, 0&, 0, 0)
    End If
End Sub

Public Sub openflashmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyD, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyD, 0&)
    Pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", "Incoming/Saved Mail") <> 0&
     lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", "Incoming/Saved Mail"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.x&, cursorpos.Y&)
End Sub

Public Sub openflashmailitem(lngindex As Long)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    lngaol& = FindWindow("AOL Frame25", vbNullString)
    lngmdi& = FindWindowEx(lngaol&, 0, "MDIClient", vbNullString)
    lngmailbox& = FindWindowEx(lngmdi&, 0, "AOL Child", "Incoming/Saved Mail")
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
        If lngindex& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, lngindex&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, lngindex&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub

Public Sub opennewmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyR, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyR, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyR, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyR, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s Online Mailbox") <> 0&
     lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s Online Mailbox"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.x, cursorpos.Y)
End Sub

Public Sub opennewmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngcount As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub


Public Sub openoldmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyO, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyO, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyO, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyO, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s Online Mailbox") <> 0&
     lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s Online Mailbox"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.x, cursorpos.Y)
End Sub

Public Sub openoldmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngcount As Long, lngindex As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub openprofile(strperson As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim mousecur As POINTAPI, lngprofwin As Long, lngedit As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(mousecur.x, mousecur.Y)
    Do: DoEvents
         lngprofwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Get a Member's Profile")
         lngedit& = FindWindowEx(lngprofwin&, 0&, "_AOL_Edit", vbNullString)
         lngicon& = FindWindowEx(lngprofwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngprofwin& <> 0& And lngedit& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strperson$)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngprofwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub opensentmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyS, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyS, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyS, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyS, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s Online Mailbox") <> 0&
     lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", GetUser$ & "'s Online Mailbox"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.x, cursorpos.Y)
End Sub
Public Sub opensentmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngcount As Long, lngindex As Long
     lngmailbox& = FindMailBox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
         lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
         lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub


Public Sub PlayCD(track$)
     Dim lRet As Long
     Dim nCurrentTrack As Integer

     'Open the device
     lRet = mciSendString("open cdaudio alias cd wait", 0&, 0, 0)

     'Set the time format to Tracks (default is milliseconds)
     lRet = mciSendString("set cd time format tmsf", 0&, 0, 0)

     'Then to play from the beginning
     lRet = mciSendString("play cd", 0&, 0, 0)

     'Or to play from a specific track, say track 4
     nCurrentTrack = track
     lRet = mciSendString("play cd from" & Str(nCurrentTrack), 0&, 0, 0)

     End Sub
Public Sub printtext(Text As String)
    Dim lngoldcursor As Long
     lngoldcursor& = Screen.MousePointer
     Screen.MousePointer = 11&
    Printer.Print (Text$)
    Printer.NewPage
    Printer.EndDoc
     Screen.MousePointer = lngoldcursor&
End Sub

Public Sub PrivateRoom(strroom As String)
    Call Keyword("aol://2719:2-2-" & strroom$)
End Sub

Public Sub PublicRoom(strroom As String)
    Call Keyword("aol://2719:21-2-" & strroom$)
End Sub

Public Function readini(strsection As String, strkey As String, strfullpath As String) As String
   Dim strBuffer As String
    strBuffer$ = String$(750, Chr$(0&))
    readini$ = Left$(strBuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strBuffer, Len(strBuffer), strfullpath$))
End Function

Public Sub removelistitem(TheList As Listbox, strentry As String)
    On Error Resume Next:
    Dim lngindex As Long
    If TheList.ListCount = 0& Then Exit Sub
    For lngindex& = 0& To TheList.ListCount - 1&
         TheList.ListIndex = lngindex&
        If LCase$(TheList.List(lngindex&)) = LCase$(strentry$) Then
            Call TheList.RemoveItem(lngindex&)
            Exit Sub
            If Err Then Exit Sub
        End If
    Next lngindex&
End Sub

Public Sub removenetzerobanner(strdirectory As String)
    If Right$(strdirectory$, 1&) <> "\" Then
         strdirectory$ = strdirectory$ & "\"
    End If
    If Dir(strdirectory$) = "" Then Exit Sub
     strdirectory$ = strdirectory$ & "bin\"
    Call Kill(strdirectory$ & "jdbcodbc.dll")
    Call Kill(strdirectory$ & "jpeg.dll")
    Call Kill(strdirectory$ & "jre.exe")
    Call Kill(strdirectory$ & "jrew.exe")
    Call Kill(strdirectory$ & "math.dll")
    DoEvents
    Call Kill(strdirectory$ & "mmedia.dll")
    Call Kill(strdirectory$ & "net.dll")
    Call Kill(strdirectory$ & "rmiregistry.exe")
    Call Kill(strdirectory$ & "symcjit.dll")
    Call Kill(strdirectory$ & "sysresource.dll")
End Sub


Public Sub roombuster(strroomname As String)
    'this is not the code i use in ghetto star room buster
    Dim lngroom As Long, lngokwin As Long, lngokbut As Long
     lngroom& = FindRoom&
    If ReplaceString$(LCase$(GetText$(lngroom&)), " ", "") = ReplaceString$(LCase$(strroomname$), " ", "") Then Exit Sub
    If lngroom& Then
        Call windowclose(lngroom&)
        Do: DoEvents
            Call Keyword("aol://2719:2-2-" & strroomname$)
            Do: DoEvents
                 lngokwin& = FindWindow("#32770", "america online")
                 lngokbut& = FindWindowEx(lngokwin&, 0&, vbNullString, "ok")
                 lngroom& = FindRoom&
            Loop Until lngokbut& Or lngroom&
            If lngokwin& <> 0& Then
                Do: DoEvents
                    Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
                    Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
                Loop Until lngokwin& = 0&
            End If
        Loop Until lngroom& <> 0&
    End If
End Sub

Public Function RoomCount() As Long
    Dim lngroom As Long, lnglist As Long, lngcount As Long
     lngroom& = FindRoom&
    If lngroom& = 0& Then
        Exit Function
    Else
         lnglist& = FindWindowEx(lngroom&, 0&, "_AOL_Listbox", vbNullString)
         lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
         RoomCount& = lngcount&
    End If
End Function

Public Sub roomrunner(strroomname As String)
    'beta testers have asked if this is the code
    'from my prog: ghetto star room buster, the
    'answer is no
    Dim lngroom As Long, lngnumber As Long, lngokwin As Long, lngokbut As Long
     lngroom& = FindRoom&
     lngnumber& = 0&
    If lngroom& Then
        windowclose (lngroom&)
        Do: DoEvents
            If lngnumber& = 0& Then lngnumber& = ""
            Call Keyword("aol://2719:2-2-" & strroomname$ & lngnumber&)
            Do: DoEvents
                 lngokwin& = FindWindow("#32770", "america online")
                 lngokbut& = FindWindowEx(lngokwin&, 0&, vbNullString, "ok")
                 lngroom& = FindRoom&
            Loop Until lngokbut& Or lngroom&
            If lngokwin& <> 0& Then
                Do
                    Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
                    Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
                Loop Until lngokwin& = 0&
            End If
            If lngnumber& = "" Then
                 lngnumber& = 1&
            Else
                 lngnumber& = lngnumber& + 1&
            End If
        Loop Until lngroom& <> 0&
    End If
End Sub

Public Sub runanymenu(lngwindow&, lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngmenu As Long, lngsubmenu As Long, lngid As Long
     lngmenu& = GetMenu(lngwindow&)
     lngsubmenu& = GetSubMenu(lngmenu&, lngmenunumber&)
     lngid& = GetMenuItemID(lngsubmenu&, lngsubmenunumber&)
    Call SendMessageLong(lngwindow&, WM_COMMAND, ByVal lngid&, 0&)
End Sub

Public Sub runanymenubystring(lngwindow As Long, strmenutext As String)
    Dim lngmmenu As Long, lngmmcount As Long, lngindex As Long
    Dim lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
     lngmmenu& = GetMenu(lngwindow&)
     lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
         lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
         lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
             lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
             strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If LCase$(strstring$) = ReplaceString(LCase$(strmenutext$), Chr$(0&), "") Then
                Call SendMessageLong(lngwindow&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Public Sub runaolmenu(lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngaol As Long, lngmenu As Long, lngsubmenu As Long, lngid As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmenu& = GetMenu(lngaol&)
     lngsubmenu& = GetSubMenu(lngmenu&, lngmenunumber&)
     lngid& = GetMenuItemID(lngsubmenu&, lngsubmenunumber&)
    Call SendMessageLong(lngaol&, WM_COMMAND, ByVal lngid&, 0&)
End Sub

Public Sub runaolmenubystring(strmenutext As String)
'ThorBrowserWndClass
    Dim lngaol As Long, lngmmenu As Long, lngmmcount As Long
    Dim lngindex As Long, lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
     lngaol& = FindWindow("aol frame25", vbNullString)
     lngmmenu& = GetMenu(lngaol&)
     lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
         lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
         lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
             lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
             strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If InStr(LCase$(strstring$), LCase$(strmenutext$)) Then
                Call SendMessageLong(lngaol&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Public Sub runoutlookmenu(strmenutext As String)
    Dim lngaol As Long, lngmmenu As Long, lngmmcount As Long
    Dim lngindex As Long, lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
     lngaol& = FindWindow("ThorBrowserWndClass", vbNullString)
     lngmmenu& = GetMenu(lngaol&)
     lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
         lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
         lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
             lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
             strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If InStr(LCase$(strstring$), LCase$(strmenutext$)) Then
                Call SendMessageLong(lngaol&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Public Sub runoutlookmenunote(strmenutext As String)
    Dim lngaol As Long, lngmmenu As Long, lngmmcount As Long
    Dim lngindex As Long, lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
     lngaol& = FindWindow("ATH_Note", vbNullString)
     lngmmenu& = GetMenu(lngaol&)
     lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
         lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
         lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
             lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
             strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If InStr(LCase$(strstring$), LCase$(strmenutext$)) Then
                Call SendMessageLong(lngaol&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Public Sub SaveComboBox(strnameandpath As String, TheCombo As ComboBox)
    Dim index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For index& = 0& To TheCombo.ListCount - 1&
        Print #1&, TheCombo.List(index&)
    Next index&
    Close #1&
End Sub

Public Sub savecomboboxwithcommondialog(commondialog32 As Control, TheCombo As ComboBox)
    Dim index As Long
    On Error Resume Next
    With commondialog32
        .DialogTitle = "load a listbox"
        .CancelError = True
        .FilterIndex = 0&
        .ShowOpen
    End With
    Open commondialog32.FileName For Output As #1&
    For index& = 0& To TheCombo.ListCount - 1&
        Print #1&, TheCombo.List(index&)
    Next index&
    Close #1&
End Sub
Public Sub SaveListBox(strnameandpath As String, TheList As Listbox)
    Dim index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For index& = 0& To TheList.ListCount - 1&
        Print #1&, TheList.List(index&)
    Next index&
    Close #1&
End Sub

Public Sub savelistboxwithcommondialog(commondialog32 As Control, TheList As Listbox)
    Dim index As Long
    On Error Resume Next
    With commondialog32
        .DialogTitle = "save a listbox"
        .CancelError = True
        .FilterIndex = 0&
        .ShowOpen
    End With
    Open commondialog32.FileName For Output As #1&
        For index& = 0& To TheList.ListCount - 1&
            Print #1&, TheList.List(index&)
        Next index&
    Close #1&
End Sub

Public Sub savesnlist(TheList As Control, commondlg32 As Control)
    Dim strsnhold As String, index As Long
    With commondlg32
        .CancelError = True
        .DialogTitle = "save a sn list"
        .Filter = "text files (*.txt)|*.txt"
        .FilterIndex = 0&
        .ShowSave
    End With
    For index& = 0& To TheList.ListCount - 1
        If index& = 0& Then
             strsnhold$ = TheList.List(index&)
        Else
             strsnhold$ = strsnhold$ & "," & TheList.List(index&)
        End If
    Next index&
    Open commondlg32.FileName For Output As #1
        Print #1, strsnhold$
    Close #1
End Sub

Public Sub scrollmacro(strText As String)
    Dim lngcounter As Long
    If Mid(strText$, Len(strText$), 1&) <> Chr$(10&) Then strText$ = strText$ & Chr$(13&) & Chr$(10&)
    Do While InStr(strText$, Chr$(13&)) <> 0&
         lngcounter& = lngcounter& + 1&
        Call sendchat(Mid$(strText$, 1&, InStr(strText$, Chr$(13&)) - 1&))
        If lngcounter = 3& Then
            Call Pause(0.6)
             lngcounter& = 0&
        End If
         strText$ = Mid$(strText$, InStr(strText$, Chr$(13&) & Chr$(10&)) + 2&)
    Loop
End Sub


Public Sub scrollprofile(strscreenname As String)
    Dim index As Long, strText As String
     strText$ = getprofile(strscreenname$)
    Call sendchat2("scrolling " & strscreenname$ & "'s profile")
    For index& = 0& To getstringlinecount(strText$)
        Call sendchat(getlinefromstring(strText$, index&))
        Pause 0.6
    Next index&
End Sub

Public Sub sendchat(strmessage As String)
    Dim lngchat As Long, lngrich As Long, strText As String
     lngchat& = FindRoom&
     lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
     lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
     strText$ = GetText(lngrich&)
    If strText$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until GetText$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strText$)
End Sub

Public Sub sendchat2(strmessage As String)
    Dim lngchat As Long, lngrich As Long, strText As String
     lngchat& = FindRoom&
     lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
     lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
     strText$ = GetText(lngrich&)
    If strText$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "÷•.• " & strmessage$)
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until GetText$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strText$)
End Sub

Public Sub sendchatlist(thecontrol As Control)
    Dim lngindex As Long
    For lngindex& = 0& To thecontrol.ListCount
        Call sendchat(thecontrol.List(lngindex&))
        Pause 0.7
    Next lngindex&
End Sub

Public Function sendchatspiral(strstring As String)
    Dim lngindex As Long, lnglen As Long
    lnglen& = Len(strstring$)
    Do: DoEvents
        lngindex& = lngindex& + 1&
        Call sendchat(Left$(strstring$, lngindex&))
        Pause 0.6
    Loop Until lngindex& = lnglen&
End Function

Public Sub SendIM(strperson As String, strmessage As String)
    Dim lngaol As Long, lngmdi As Long, lngimwin As Long, lngsendbut As Long
    Dim lngrich As Long, lngok As Long, lngokbut As Long, lngokwin2 As Long
    Dim lngokwin As Long, lngokbut2 As Long, lngokwin3 As Long, lngsendbut1 As Long
    Dim lngsendbut2 As Long, lngsendbut3 As Long, lngsendbut4 As Long
    Dim lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long
    Call Keyword("aol://9293:" & strperson$)
    Do: DoEvents
         lngimwin& = FindWindowEx(lngmdi&, 0&, "aol child", "send instant message")
         lngrich& = FindWindowEx(lngimwin&, 0&, "richcntl", vbNullString)
    Loop Until lngimwin& <> 0& And lngrich& <> 0&
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Do: DoEvents: Loop Until GetText(lngrich&) = strmessage$
    Do: DoEvents
         lngsendbut& = FindWindowEx(lngimwin&, 0&, "_aol_icon", vbNullString)
         lngsendbut1& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut2& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut3& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut4& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut5& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut6& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut7& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
         lngsendbut& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
    Loop Until lngsendbut& <> 0& And lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And lngsendbut7& <> lngsendbut&
     lngimwin& = FindWindowEx(lngmdi&, 0&, "aol child", "send instant message")
     lngokwin& = FindWindow("#32770", "america online")
     lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "ok")
     lngokwin2& = FindWindowEx(lngaol&, 0&, "#32770", vbNullString)
     lngokbut2& = FindWindowEx(lngokwin2&, 0&, "button", vbNullString)
    Do: DoEvents
         lngokwin2& = FindWindowEx(lngaol&, 0&, "#32770", vbNullString)
         lngokbut2& = FindWindowEx(lngokwin2&, 0&, "button", vbNullString)
        Call SendMessageLong(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        If LCase$(strperson$) = "$im_off" Then
            Do: DoEvents: Loop Until lngokwin2& <> 0&
        End If
        If LCase$(strperson$) = "$im_on" Then
            Do: DoEvents: Loop Until lngokwin2& <> 0&
        End If
    Loop Until lngsendbut& = 0& Or lngokwin2& <> 0& Or lngokwin& <> 0&
    If lngokwin2& <> 0& Then
        Do: DoEvents
            Call SendMessageLong(lngokwin2&, WM_CLOSE, 0&, 0&)
            Call SendMessageLong(lngimwin&, WM_CLOSE, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
            Pause 0.7
        Loop Until lngokwin2& = 0&
    End If
    Do: DoEvents
         lngokwin3& = FindWindowEx(lngaol&, lngokwin2&, "#32770", vbNullString)
         lngokbut2& = FindWindowEx(lngokwin3&, 0&, "button", vbNullString)
        Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessageLong(lngokwin&, WM_CLOSE, 0&, 0&)
        Call SendMessageLong(lngimwin&, WM_CLOSE, 0&, 0&)
    Loop Until lngokwin& = 0&
    If lngimwin& <> 0& Then Call SendMessageLong(lngimwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub sendimresponse(strmessage As String)
    Dim lngimwin As Long, lngrich As Long, lngicon As Long
    Dim lngicon1 As Long, lngicon2 As Long, lngicon3 As Long, lngicon4 As Long
    Dim lngicon5 As Long, lngicon6 As Long, lngicon7 As Long, lngicon8 As Long
     lngimwin& = FindIM&
    If lngimwin& = 0& Then Exit Sub
     lngrich& = FindWindowEx(lngimwin&, 0&, "RICHCNTL", vbNullString)
     lngrich& = FindWindowEx(lngimwin&, lngrich&, "RICHCNTL", vbNullString)
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
     lngicon1& = FindWindowEx(lngimwin&, 0&, "_AOL_Icon", vbNullString)
     lngicon2& = FindWindowEx(lngimwin&, lngicon1&, "_AOL_Icon", vbNullString)
     lngicon3& = FindWindowEx(lngimwin&, lngicon2&, "_AOL_Icon", vbNullString)
     lngicon4& = FindWindowEx(lngimwin&, lngicon3&, "_AOL_Icon", vbNullString)
     lngicon5& = FindWindowEx(lngimwin&, lngicon4&, "_AOL_Icon", vbNullString)
     lngicon6& = FindWindowEx(lngimwin&, lngicon5&, "_AOL_Icon", vbNullString)
     lngicon7& = FindWindowEx(lngimwin&, lngicon6&, "_AOL_Icon", vbNullString)
     lngicon8& = FindWindowEx(lngimwin&, lngicon7&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngimwin&, lngicon8&, "_AOL_Icon", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub SendMail(Person As String, SUBJECT As String, Message As String)
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, SUBJECT$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub


Public Sub sendmailattachment(strperson, strsubject As String, strmessage As String, strfilepath As String, strfilename As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbarwin As Long, lngtoolbarico As Long, lngcomposeico As Long, lngokbut As Long
    Dim lngmailwin As Long, lngeditwin As Long, lngsubjectwin As Long, lngrich As Long, lngattachbut As Long
    Dim lngsendbut As Long, lngsendbut1 As Long, lngsendbut2 As Long, lngsendbut3 As Long, lngsendbut4 As Long
    Dim lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long, lngsendbut9 As Long
    Dim lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long, lngattachwin As Long
    Dim lngbrowsewin As Long, lngedit As Long, lngopenbut As Long, lngcombo As Long
     lngaol& = FindWindow("aol frame25", "america  online")
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If lngmdi& <= 0& Then
        Exit Sub
    Else
         lngtoolbarwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
         lngtoolbarico& = FindWindowEx(lngtoolbarwin&, 0&, "_aol_toolbar", vbNullString)
         lngcomposeico& = FindWindowEx(lngtoolbarico&, 0&, "_aol_icon", vbNullString)
         lngcomposeico& = FindWindowEx(lngtoolbarico&, lngcomposeico&, "_aol_icon", vbNullString)
        Call SendMessageLong(lngcomposeico&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngcomposeico&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngmailwin& = FindWindowEx(lngmdi&, 0&, "aol child", "write mail")
             lngeditwin& = FindWindowEx(lngmailwin&, 0&, "_aol_edit", vbNullString)
        Loop Until lngmailwin& <> 0& And lngeditwin& <> 0&
        Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strperson)
        Do: DoEvents
             lngeditwin& = FindWindowEx(lngmailwin&, lngeditwin&, "_aol_edit", vbNullString)
             lngeditwin& = FindWindowEx(lngmailwin&, lngeditwin&, "_aol_edit", vbNullString)
             lngrich& = FindWindowEx(lngmailwin&, 0&, "richcntl", vbNullString)
        Loop Until lngmailwin& <> 0& And lngeditwin& <> 0& And lngrich& <> 0&
        Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strsubject$)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
        Do: DoEvents
             lngsendbut1& = FindWindowEx(lngmailwin&, 0&, "_aol_icon", vbNullString)
             lngsendbut2& = FindWindowEx(lngmailwin&, lngsendbut1&, "_aol_icon", vbNullString)
             lngsendbut3& = FindWindowEx(lngmailwin&, lngsendbut2&, "_aol_icon", vbNullString)
             lngsendbut4& = FindWindowEx(lngmailwin&, lngsendbut3&, "_aol_icon", vbNullString)
             lngsendbut5& = FindWindowEx(lngmailwin&, lngsendbut4&, "_aol_icon", vbNullString)
             lngsendbut6& = FindWindowEx(lngmailwin&, lngsendbut5&, "_aol_icon", vbNullString)
             lngsendbut7& = FindWindowEx(lngmailwin&, lngsendbut6&, "_aol_icon", vbNullString)
             lngsendbut8& = FindWindowEx(lngmailwin&, lngsendbut7&, "_aol_icon", vbNullString)
             lngsendbut9& = FindWindowEx(lngmailwin&, lngsendbut8&, "_aol_icon", vbNullString)
             lngsendbut10& = FindWindowEx(lngmailwin&, lngsendbut9&, "_aol_icon", vbNullString)
             lngsendbut11& = FindWindowEx(lngmailwin&, lngsendbut10&, "_aol_icon", vbNullString)
             lngsendbut12& = FindWindowEx(lngmailwin&, lngsendbut11&, "_aol_icon", vbNullString)
             lngsendbut13& = FindWindowEx(lngmailwin&, lngsendbut12&, "_aol_icon", vbNullString)
             lngsendbut& = FindWindowEx(lngmailwin&, lngsendbut13&, "_aol_icon", vbNullString)
        Loop Until lngsendbut& <> 0& And _
        lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
        lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
        lngsendbut13& <> lngsendbut&
         lngattachbut& = lngsendbut13&
        Call PostMessage(lngattachbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngattachbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngattachwin& = FindWindow("_AOL_Modal", "Attachments")
             lngattachbut& = FindWindowEx(lngattachwin&, 0&, "_AOL_Icon", vbNullString)
             lngokbut& = FindWindowEx(lngattachwin&, lngattachbut&, "_AOL_Icon", vbNullString)
             lngokbut& = FindWindowEx(lngattachwin&, lngokbut&, "_AOL_Icon", vbNullString)
        Loop Until lngattachwin& <> 0& And lngattachbut& <> 0& And lngokbut& <> 0&
        Call PostMessage(lngattachbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngattachbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngbrowsewin& = FindWindow("#32770", "Attach")
             lngedit& = FindWindowEx(lngbrowsewin&, 0&, "Edit", vbNullString)
             lngopenbut& = FindWindowEx(lngbrowsewin&, 0&, "Button", "&Open")
             lngcombo& = FindWindowEx(lngbrowsewin&, 0&, "ComboBox", vbNullString)
        Loop Until lngbrowsewin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfilepath$)
        Call PostMessage(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngcombo& = FindWindowEx(lngbrowsewin&, 0&, "ComboBox", vbNullString)
        Loop Until GetText(lngcombo&) = strfilepath$
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfilename$)
        Do: DoEvents
             lngbrowsewin& = FindWindow("#32770", "Attach")
             lngopenbut& = FindWindowEx(lngbrowsewin&, 0&, "Button", "&Open")
            Call PostMessage(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
            Pause 0.2
        Loop Until lngbrowsewin& = 0&
        Do: DoEvents
             lngattachwin& = FindWindow("_AOL_Modal", "Attachments")
            Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
            Pause 0.2
        Loop Until lngattachwin& = 0&
        Do While lngmailwin& <> 0&
            Call SendMessageLong(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
             lngmailwin& = FindWindowEx(lngmdi&, 0&, "aol child", "write mail")
            Pause 0.2
        Loop: DoEvents
    End If
End Sub

Public Sub sendmailbcc(strpeople As String, strsubject As String, strmessage As String)
    Call SendMail("(" & GetUser$ & "," & strpeople$ & ")", strsubject$, strmessage$)
End Sub


Public Sub sendmassimwithdogbar(strwho As Control, strmessage As String, progbar As Control)
    Dim lngindex As Long
     progbar.Max = strwho.ListCount - 1&
    For lngindex& = 0& To strwho.ListCount - 1&
         progbar.Value = lngindex& - 1&
        Call SendIM(strwho.List(lngindex&), strmessage$)
        Pause 0.7
    Next lngindex&
End Sub

Public Sub sendmassmailfast(strwho As Control, strsubject As String, strmessage As String)
    'this is a fast way to do it
    Dim lngindex As Long, strpeople As String
    For lngindex& = 0& To strwho.ListCount - 1&
         strpeople$ = strpeople$ & strwho.List(lngindex&) & ","
    Next lngindex&
    If Left$(strpeople$, 1&) = "," Then
         strpeople$ = Right$(strpeople$, Len(strpeople$) - 1&)
    End If
    If Right$(strpeople$, 1&) = "," Then
         strpeople$ = Left$(strpeople$, Len(strpeople$) - 1&)
    End If
    Call SendMail(strpeople$, strsubject$, strmessage$)
End Sub

Public Sub sendmassmailslow(strwho As Control, strsubject As String, strmessage As String)
    'this is a   s  l  o  w   way to do it
    Dim lngindex As Long
    For lngindex& = 0& To strwho.ListCount - 1&
        Call SendMail(strwho.List(lngindex&), strsubject$, strmessage$)
    Next lngindex&
End Sub

Public Sub sendpage(pagerid As String, Message As String)
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long, lngsendpage As Long
    Dim lngicon1 As Long, lngicon2 As Long, lngicon3 As Long, lngsendwin As Long
    Dim lngsendbut As Long, lngidbox As Long, lngmsgbox As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyE, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyE, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyE, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyE, 0&)
    Do: DoEvents
         lngsendpage& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Send Page")
         lngicon1& = FindWindowEx(lngsendpage&, 0&, "_AOL_Icon", vbNullString)
         lngicon2& = FindWindowEx(lngsendpage&, lngicon1&, "_AOL_Icon", vbNullString)
         lngicon3& = FindWindowEx(lngsendpage&, lngicon2&, "_AOL_Icon", vbNullString)
         lngicon& = FindWindowEx(lngsendpage&, lngicon3&, "_AOL_Icon", vbNullString)
    Loop Until lngsendpage& <> 0& And lngicon1 <> 0& And lngicon2 <> 0& And lngicon3 <> 0& And lngicon <> 0& And lngicon1 <> lngicon2& And lngicon2& <> lngicon3& And lngicon3& <> lngicon&
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
         lngsendwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", ReplaceString(UCase$(GetUser$), " ", "") & "'s Send a Page")
         lngidbox& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
         lngmsgbox& = FindWindowEx(lngsendwin&, lngidbox&, "_AOL_Edit", vbNullString)
         lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngsendwin& <> 0& And lngidbox& <> 0& And lngmsgbox& <> 0& And lngicon& <> 0&: Pause 0.2
    Call SendMessage(lngsendpage&, WM_CLOSE, 0&, 0&)
    Call SendMessageByString(lngidbox&, WM_SETTEXT, 0&, pagerid$)
    Call SendMessageByString(lngmsgbox&, WM_SETTEXT, 0&, Message$)
    Call PostMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
    Pause 4
    Call SendMessage(lngsendwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub sendtext(lngwindow As Long, strText As String)
    Call SendMessageByString(lngwindow&, WM_SETTEXT, 0&, strText$)
End Sub

Public Function servercreatelist(TheList As Control, strseperator As String) As String
    Dim index As Long, strmailstring As String
    For index& = 0& To TheList.ListCount - 1&
         strmailstring$ = strmailstring$ & "~" & index& & "~ " & TheList.List(index&) & vbCrLf$
    Next index&
     servercreatelist$ = strmailstring$
End Function

Public Function serverlookfor(TheList As Control, strentry As String) As String
    Dim index As Long
    If TheList.ListCount = 0& Then Exit Function
    For index& = 0& To TheList.ListCount - 1&
        If InStr(LCase$(TheList.List(index&)), strentry$) > -1& Then
             serverlookfor$ = serverlookfor$ & vbCrLf$ & TheList.List(index&)
            If Err Then Exit Function
        End If
    Next index&
End Function

Public Sub serversendfind(strwho As String, strrequest As String, TheList As Listbox)
    Dim endstring As String, findlist As String
     findlist$ = serverlookfor(TheList, strrequest$)
    If findlist$ = "" Then
        If strrequest$ = "my penis" Then strrequest$ = "your penis"
        Call sendchat("~•~ sorry " & strwho$ & ", but " & strrequest & " wasn't found")
        Exit Sub
    End If
    Call SendMail(strwho$, "~•~ find results for " & strrequest$ & " ~•~", "please remember that you requested this mail and that " & GetUser$ & ", does nottake responsibility for what you do with the contents of this mail, this is for entertainment purposes only." & vbCrLf$ & stringlinked("http://www.nkillaz.com/izekial/", "izekial32.bas, click here") & vbCrLf$ & findlist$)
    Call sendchat("~•~ " & strwho$ & ", the results for " & strrequest$ & " were sent")
End Sub

Public Sub SetMailPrefs()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim lngprefwin As Long, lngbut As Long, lnggenstatic As Long, lngmailstatic As Long
    Dim lngfontstatic As Long, lngmarketstatic As Long, lngmodal As Long, chkconfirm As Long
    Dim chkclose As Long, chkspellchk As Long, lngokbut As Long, cursorpos As POINTAPI
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyP, 0&)
    Call SetCursorPos(cursorpos.x, cursorpos.Y)
    Do: DoEvents
         lngmodal& = FindWindow("_AOL_Modal", "Mail Preferences")
         lngokbut& = FindWindowEx(lngmodal&, 0&, "_AOL_icon", vbNullString)
        Pause 1
    Loop Until lngmodal& <> 0& And lngokbut& <> 0&
     chkconfirm& = FindWindowEx(lngmodal&, 0&, "_AOL_Checkbox", vbNullString)
     chkclose& = FindWindowEx(lngmodal&, chkconfirm&, "_AOL_Checkbox", vbNullString)
     chkspellchk& = FindWindowEx(lngmodal&, chkclose&, "_AOL_Checkbox", vbNullString)
     chkspellchk& = FindWindowEx(lngmodal&, chkspellchk&, "_AOL_Checkbox", vbNullString)
     chkspellchk& = FindWindowEx(lngmodal&, chkspellchk&, "_AOL_Checkbox", vbNullString)
     chkspellchk& = FindWindowEx(lngmodal&, chkspellchk&, "_AOL_Checkbox", vbNullString)
    Call SendMessage(chkconfirm&, BM_SETCHECK, False, 0&)
    Call SendMessage(chkclose&, BM_SETCHECK, True, 0&)
    Call SendMessage(chkspellchk&, BM_SETCHECK, False, 0&)
    Call SendMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        lngokbut& = FindWindowEx(lngmodal&, 0&, "_AOL_icon", vbNullString)
    Loop Until lngokbut& = 0&
    Call PostMessage(lngprefwin&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub setprofile(strName As String, strlocation As String, male As Boolean, female As Boolean, nosex As Boolean, strlovestatus As String, strhobbies As String, strcomputer As String, stroccupation As String, strquote As String)
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, mousecur As POINTAPI
    Dim lngmenu As Long, lngmdi As Long, lngprofwin As Long, lngname As Long
    Dim lngmsex As Long, lngfsex As Long, lngnosex As Long, lngupbut As Long
    Dim lnglocation As Long, lngbirthday As Long, lngmstatus As Long, lnghobbies As Long
    Dim computer As Long, lngother As Long, lngquote As Long, lngcomputer As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyY, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyY, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyY, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyY, 0&)
    Call SetCursorPos(mousecur.x, mousecur.Y)
    Do: DoEvents
         lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
         lngprofwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Edit Your Online Profile")
         lngname& = FindWindowEx(lngprofwin&, 0&, "_AOL_Edit", vbNullString)
         lnglocation& = FindWindowEx(lngprofwin&, lngname&, "_AOL_Edit", vbNullString)
         lngmstatus& = FindWindowEx(lngprofwin&, lnglocation&, "_AOL_Edit", vbNullString)
         lnghobbies& = FindWindowEx(lngprofwin&, lngmstatus&, "_AOL_Edit", vbNullString)
         lngcomputer& = FindWindowEx(lngprofwin&, lnghobbies&, "_AOL_Edit", vbNullString)
         lngother& = FindWindowEx(lngprofwin&, lngcomputer&, "_AOL_Edit", vbNullString)
         lngquote& = FindWindowEx(lngprofwin&, lngother&, "_AOL_Edit", vbNullString)
         lngmsex& = FindWindowEx(lngprofwin&, 0&, "_AOL_Checkbox", vbNullString)
         lngfsex& = FindWindowEx(lngprofwin&, lngmsex&, "_AOL_Checkbox", vbNullString)
         lngnosex& = FindWindowEx(lngprofwin&, lngfsex&, "_AOL_Checkbox", vbNullString)
         lngupbut& = FindWindowEx(lngprofwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngprofwin& <> 0& And lngquote& <> 0& And lngnosex& <> 0& And lngupbut& <> 0&
    Call SendMessageByString(lngname&, WM_SETTEXT, 0&, strName$): DoEvents
    Call SendMessageByString(lnglocation&, WM_SETTEXT, 0&, strlocation$): DoEvents
    Call SendMessageByString(lngmstatus&, WM_SETTEXT, 0&, strlovestatus$): DoEvents
    Call SendMessageByString(lnghobbies&, WM_SETTEXT, 0&, strhobbies$): DoEvents
    Call SendMessageByString(lngcomputer&, WM_SETTEXT, 0&, strcomputer$): DoEvents
    Call SendMessageByString(lngother&, WM_SETTEXT, 0&, stroccupation$): DoEvents
    Call SendMessageByString(lngquote&, WM_SETTEXT, 0&, strquote$): DoEvents
    If male = True Then
        Call PostMessage(lngmsex&, BM_SETCHECK, True, 0&)
    ElseIf female = True Then
        Call PostMessage(lngfsex&, BM_SETCHECK, True, 0&)
    ElseIf nosex = True Then
        Call PostMessage(lngnosex&, BM_SETCHECK, True, 0&)
    End If
    Call PostMessage(lngupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngupbut&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
End Sub

Public Sub signoff()
    Call runaolmenubystring("&Sign Off")
End Sub

Public Sub signonguest(strscreenname As String, strpassword As String)
    Dim lngaol As Long, lngmdi As Long, lngsignonwin As Long, lngcombo As Long
    Dim lngedit As Long, lngsignonbut1 As Long, lngsignonbut2 As Long, lngsignonbut3 As Long
    Dim lngsignonbut As Long, lngedit1 As Long, lngokbut As Long, lngmodal As Long
    Dim lngokbut1 As Long, lngokbut2 As Long, lngerrwin As Long, lngsignonwin2 As Long
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     lngsignonwin& = findsignonwin&
    If lngsignonwin& = 0& Then
        Call runaolmenubystring("&Sign On Screen")
    End If
    Do: DoEvents
         lngsignonwin& = findsignonwin&
         lngsignonwin2& = findsignoffwin&
    Loop Until (lngsignonwin& <> 0& Or lngsignonwin2& <> 0&)
    If lngsignonwin& = 0& And lngsignonwin2& <> 0& Then
         lngsignonwin& = lngsignonwin2&
    End If
    Do: DoEvents
         lngcombo& = FindWindowEx(lngsignonwin&, 0&, "_AOL_Combobox", vbNullString)
         lngedit& = FindWindowEx(lngsignonwin&, 0&, "_AOL_Edit", vbNullString)
         lngsignonbut1& = FindWindowEx(lngsignonwin&, 0&, "_AOL_Icon", vbNullString)
         lngsignonbut2& = FindWindowEx(lngsignonwin&, lngsignonbut1&, "_AOL_Icon", vbNullString)
         lngsignonbut3& = FindWindowEx(lngsignonwin&, lngsignonbut2&, "_AOL_Icon", vbNullString)
         lngsignonbut& = FindWindowEx(lngsignonwin&, lngsignonbut3&, "_AOL_Icon", vbNullString)
    Loop Until lngcombo& <> 0& And lngsignonbut1& <> 0& And lngsignonbut2& <> 0& And lngsignonbut3& <> 0& And lngsignonbut& <> 0& And _
    lngsignonbut1& <> lngsignonbut2& And lngsignonbut2& <> lngsignonbut3& And lngsignonbut3& <> lngsignonbut&
    Call PostMessage(lngcombo&, CB_SETCURSEL, SendMessage(lngcombo&, CB_GETCOUNT, 0&, 0&) - 1&, 0&)
    If lngedit& <> 0& Then
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, "yo b")
    End If
    Call PostMessage(lngsignonbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngsignonbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
         lngmodal& = FindWindow("_AOL_Modal", vbNullString)
         lngedit1& = FindWindowEx(lngmodal&, 0&, "_AOL_Edit", vbNullString)
         lngedit& = FindWindowEx(lngmodal&, lngedit1&, "_AOL_Edit", vbNullString)
         lngokbut1& = FindWindowEx(lngmodal&, 0&, "_AOL_Icon", vbNullString)
         lngokbut& = FindWindowEx(lngmodal&, lngokbut1&, "_AOL_Icon", vbNullString)
    Loop Until lngmodal& <> 0& And lngedit1& <> 0& And lngedit& <> 0& And lngokbut1& <> 0& And lngokbut& <> 0&: Pause 0.2
    Call SendMessageByString(lngedit1&, WM_SETTEXT, 0&, strscreenname$)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strpassword$)
    Do: DoEvents
        Call PostMessage(lngokbut1&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngokbut1&, WM_LBUTTONUP, 0&, 0&)
         lngmodal& = FindWindow("_AOL_Modal", vbNullString)
         lngerrwin& = FindWindow("#32770", "America Online")
     Pause (1)
    Loop Until lngmodal& = 0& Or lngerrwin& <> 0&: Pause 0.2
    If lngerrwin& <> 0& Then
         lngokbut2& = FindWindowEx(lngerrwin&, 0&, "Button", vbNullString)
        Do: DoEvents
            Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
            Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
        Pause (1)
        Loop Until findsignonwin& <> 0& Or findwelcomewin&
    End If
End Sub

Public Sub startdownloadlater()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Dim lngdlmanager As Long, lngicon1 As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
     lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
     lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
     lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
         lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyD, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyD, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyD, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyD, 0&)
    Do: DoEvents
        lngdlmanager& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Download Manager")
        lngicon1& = FindWindowEx(lngdlmanager&, 0&, "_AOL_Icon", vbNullString)
        lngicon& = FindWindowEx(lngdlmanager&, lngicon1&, "_AOL_Icon", vbNullString)
    Loop Until lngdlmanager& <> 0& And lngicon& <> 0& And lngicon1 <> 0& And lngicon1& <> lngicon&
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngdlmanager&, WM_CLOSE, 0&, 0&)
    Call SetCursorPos(cursorpos.x, cursorpos.Y)
End Sub

     Public Sub StopCD()
     Dim lRet As Long

     'Stop the playback
     lRet = mciSendString("stop cd wait", 0&, 0, 0)

     DoEvents  ' Windows process the event

     'Close the device
     lRet = mciSendString("close cd", 0&, 0, 0)

     End Sub

Public Function stringcustomized(strstring As String, streveryotherchar As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
     lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
         lngindex& = lngindex& + 1&
         strnextchar$ = Mid$(strstring$, lngindex&, 1&)
         strnextchar$ = strnextchar$ & streveryotherchar$
         stringcustomized$ = stringcustomized$ & strnextchar$
    Loop
End Function

Public Function stringdotted(strstring As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
     lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
         lngindex& = lngindex& + 1&
         strnextchar$ = Mid$(strstring$, lngindex&, 1&)
         strnextchar$ = strnextchar$ & "o"
         stringdotted$ = stringdotted$ & strnextchar$
    Loop
End Function

Public Function stringlinked(strurl As String, strText$) As String
     stringlinked$ = "< a href=" & Chr$(34&) & strurl$ & Chr$(34&) & ">" & strText$ & "</a>"
End Function


Public Function stringreverse(strstring As String) As String
    'note: for vb6 users this code can be replaced with
    'strstring$ = strreverse(strstring$)
    Dim lnglen As Long, lngindex As Long, strnextchar As String
     lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
         lngindex& = lngindex& + 1&
         strnextchar$ = Mid$(strstring$, lngindex&, 1&)
         stringreverse$ = strnextchar$ & stringreverse$
    Loop
End Function
Public Function stringspaced(strstring As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
     lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
         lngindex& = lngindex& + 1&
         strnextchar$ = Mid$(strstring$, lngindex&, 1&)
         strnextchar$ = strnextchar$ & " "
         stringspaced$ = stringspaced$ & strnextchar$
    Loop
End Function

Public Sub switchscreenname(index As Long, strpassword As String)
    Dim lngaol As Long, lngmdi As Long, lngswitchwin As Long
    Dim lnglist As Long, lngicon As Long, lngedit As Long
    Call runaolmenu(3&, 0&)
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
         lngswitchwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Switch Screen Names")
         lnglist& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Listbox", vbNullString)
         lngicon& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngswitchwin& <> 0& And lnglist& <> 0& And lngicon& <> 0&
    If index& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& Then
        Exit Sub
    Else
        Call SendMessageLong(lnglist&, LB_SETCURSEL, CLng(index&), 0&)
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngswitchwin& = FindWindow("_AOL_Modal", "Switch Screen Name")
             lngicon& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Icon", vbNullString)
        Loop Until lngswitchwin& <> 0& And lngicon& <> 0&
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
             lngswitchwin& = FindWindow("_AOL_Modal", "Switch Screen Name")
             lngedit& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Edit", vbNullString)
             lngicon& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Icon", vbNullString)
        Loop Until lngswitchwin& <> 0& And lngedit& <> 0& And lngicon& <> 0&
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strpassword$)
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    End If
End Sub

Public Function trimchar(strText As String, strChar As String) As String
    trimchar$ = ReplaceString(strText$, strChar$, "")
End Function

Public Function trimenters(strText As String) As String
     trimenters$ = ReplaceString(strText$, Chr$(13&), "")
     trimenters$ = ReplaceString(trimenters$, Chr$(10&), "")
End Function

Public Function TrimSpaces(strText As String) As String
     TrimSpaces$ = ReplaceString(strText$, " ", "")
End Function

Public Sub usernamereplace(strsn As String, strnewsn As String, straolpath As String)
    Dim lngsnlen As Long, lngnumhold As Long, strText As String, lngthesn As Long
    Dim lngthesn2 As Long, strtoreplace As String, lnglenofile2 As Long
    Call windowclose(findsignonwin&)
     lngsnlen& = Len(strsn$)
    Select Case lngsnlen&
        Case 3&
             strsn$ = strsn$ & "       "
        Case 4&
             strsn$ = strsn$ & "      "
        Case 5&
             strsn$ = strsn$ & "     "
        Case 6&
             strsn$ = strsn$ & "    "
        Case 7&
             strsn$ = strsn$ & "   "
        Case 8&
             strsn$ = strsn$ & "  "
        Case 9&
             strsn$ = strsn$ & " "
        Case 10&
             strsn$ = strsn$
    End Select
     lngsnlen& = Len(strnewsn$)
    Select Case lngsnlen&
        Case 3&
             strnewsn$ = strnewsn$ & "       "
        Case 4&
             strnewsn$ = strnewsn$ & "      "
        Case 5&
             strnewsn$ = strnewsn$ & "     "
        Case 6&
             strnewsn$ = strnewsn$ & "    "
        Case 7&
             strnewsn$ = strnewsn$ & "   "
        Case 8&
             strnewsn$ = strnewsn$ & "  "
        Case 9&
             strnewsn$ = strnewsn$ & " "
        Case 10&
             strnewsn$ = strnewsn$
    End Select
     lngnumhold& = 1&
    If Right$(straolpath$, 1&) = "\" Then straolpath$ = Right$(straolpath$, 1&)
    Do Until 2& > 3&: DoEvents
         strText$ = ""
        On Error Resume Next
        Open straolpath$ & "\idb\main.idx" For Binary As #1&
            If Err Then
                Exit Sub
            Else
                 strText$ = String$(32000&, 0&)
                Get #1, lngnumhold&, strText$
            End If
        Close #1&
        Open straolpath$ & "\idb\main.idx" For Binary As #2&
             lngthesn& = InStr(1&, strText$, strsn$, 1&)
            If lngthesn& Then
                Mid(strText$, lngthesn&) = strnewsn$
                 strtoreplace$ = strnewsn$
                Put #2&, (lngnumhold& + lngthesn&) - 1&, strtoreplace$
getthatsn:
                DoEvents
                 lngthesn2& = InStr(1&, strText$, strsn$, 1&)
                If lngthesn2& Then
                    Mid(strText$, lngthesn2&) = strnewsn$
                    Put #2&, (lngnumhold& + lngthesn2&) - 1&, strtoreplace$
                    GoTo getthatsn
                End If
            End If
             lngnumhold& = lngnumhold& + 32000&
             lnglenofile2& = LOF(2&)
        Close #2&
        If lngnumhold& > lnglenofile2& Then Call runaolmenubystring("&Sign On Screen"): Exit Sub
    Loop
End Sub

Public Sub waitforlisttoload(lnglist As Long)
    Dim lngcount As Long
    Do: DoEvents
         lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&): Call Pause(2&)
        If lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Do
    Loop
End Sub

Public Sub waitforok()
    Dim lngokwin As Long, lngokbut As Long, lngokwin2 As Long, lngokbut2 As Long
    Do: DoEvents
         lngokwin& = FindWindow("#32770", "America Online")
         lngokbut& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
         lngokwin2& = FindWindow("_AOL_Modal", "America Online")
         lngokbut2& = FindWindowEx(lngokwin2&, 0&, "_AOL_Icon", vbNullString)
    Loop Until (lngokwin& <> 0& And lngokbut& <> 0&) Or (lngokwin2& <> 0& And lngokbut2& <> 0&)
    If lngokwin& <> 0& Then
        Do: DoEvents
            Pause 0.2
            Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
             lngokwin& = FindWindow("#32770", "America Online")
             lngokbut& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Loop Until lngokwin& = 0&
    ElseIf lngokwin2& <> 0& Then
        Do: DoEvents
            Pause 0.2
            Call SendMessageLong(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
             lngokwin2& = FindWindow("_AOL_Modal", "America Online")
             lngokbut2& = FindWindowEx(lngokwin2&, 0&, "_AOL_Icon", vbNullString)
        Loop Until lngokwin2& = 0&
    End If
End Sub

Public Function waitforokorchatroom(room As String)
    Dim strroomname As String, lngokwin As Long, lngbutton As Long
     room$ = LCase$(ReplaceString(room$, " ", ""))
    Do: DoEvents
         strroomname$ = GetCaption(FindRoom&)
         strroomname$ = LCase$(ReplaceString(strroomname$, " ", ""))
         lngokwin& = FindWindow("#32770", "america online")
         lngbutton& = FindWindowEx(lngokwin&, 0&, "button", "ok")
    Loop Until (lngokwin& <> 0& And lngbutton& <> 0&) Or room$ = strroomname$
    DoEvents
    If lngokwin& <> 0& Then
        Do: DoEvents
            Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
             lngokwin& = FindWindow("#32770", "america online")
             lngbutton& = FindWindowEx(lngokwin&, 0&, "button", "ok")
        Loop Until lngokwin& = 0& And lngbutton& = 0&
    End If
End Function

Public Sub waitfortextboxtoload(lngtextbox As Long)
    Dim strText As String
    Do: DoEvents
         strText$ = GetText(lngtextbox&): Call Pause(2&)
        If strText$ = GetText(lngtextbox&) Then Exit Do
    Loop
End Sub

Public Sub wavloop(File As String)
    Dim lngflags As Long
    If file_ifileexists(File$) = False Then
        Exit Sub
    Else
         lngflags& = SND_ASYNC Or SND_LOOP
        Call sndPlaySound(File$, lngflags&)
    End If
End Sub

Public Sub wavloopstop()
    Dim lngflags As Long
     lngflags& = SND_ASYNC Or SND_LOOP
    Call sndPlaySound("", lngflags&)
End Sub

Public Sub wavstop()
    Call sndPlaySound("", SND_FLAG)
End Sub

Public Sub win98changeresolution(screenwidth As Single, screenheight As Single)
    Dim issafe As Boolean, index As Long, lngchange As Long
    Do: DoEvents
         issafe = EnumDisplaySettings(0&, index&, DevM)
         index& = index& + 1&
    Loop Until issafe = False
     DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
     DevM.dmPelsWidth = screenwidth
     DevM.dmPelsHeight = screenheight
     lngchange& = ChangeDisplaySettings(DevM, 0)
End Sub

Public Sub win98changewallpaper(strFile As String)
    Dim lngsysparam As Long
    lngsysparam& = SystemParametersInfo(ByVal 20&, vbNullString, ByVal strFile$, &H1)
    If lngsysparam& = 0& Then Exit Sub
End Sub

Public Sub win98clickstart()
    Dim lngshell As Long, lngstartbut As Long
     lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
     lngstartbut& = FindWindowEx(lngshell&, 0, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub win98controlpanel()
    Call Shell("RUNDLL32.EXE SHELL32.DLL,Control_RunDLL ,", vbNormalFocus)
End Sub

Public Sub win98findfiles(strtrigger As String)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Dim lngfindwin As Long, lngedit As Long, lngshelldef As Long
    Dim mousecur As POINTAPI, lngsyslist As Long
     lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
     lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
     lngbasebar& = FindWindow("BaseBar", vbNullString)
     lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
     lngtoolwin& = FindWindowEx(lngmenusite&, 0&, "ToolbarWindow32", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Height, Screen.Width)
    Pause 0.2
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyF, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyRight, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyRight, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
         lngfindwin& = FindWindow("#32770", "Find: All Files")
         lngedit& = FindWindowEx(lngfindwin&, 0&, "#32770", vbNullString)
         lngedit& = FindWindowEx(lngedit&, 0&, "ComboBox", vbNullString)
         lngedit& = FindWindowEx(lngedit&, 0&, "Edit", vbNullString)
         lngokbut& = FindWindowEx(lngfindwin&, 0&, "Button", "F&ind Now")
    Loop Until lngfindwin& <> 0& And lngedit& <> 0& And lngokbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strtrigger$)
    Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        lngshelldef& = FindWindowEx(lngfindwin&, 0&, "SHELLDLL_DefView", vbNullString)
        lngsyslist& = FindWindowEx(lngshelldef&, 0&, "SysListView32", vbNullString)
    Loop Until lngshelldef& <> 0& And lngsyslist& <> 0&
    Call waitforlisttoload(lngsyslist&)
    Call SetCursorPos(mousecur.x, mousecur.Y)
End Sub

Public Function getlistitemtext(lnglist As Long, ByVal index As Long) As String
    On Error Resume Next
    Dim rList As Long, sThread As Long, mThread As Long
    Dim ScreenName As String, itmHold As Long, psnHold As Long
    Dim rBytes As Long, cProcess As Long
     rList& = lnglist&
     sThread& = GetWindowThreadProcessId(rList&, cProcess&)
     mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
         ScreenName$ = String$(4, vbNullChar)
         itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
         itmHold& = itmHold& + 24
        Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes&)
        Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
         psnHold& = psnHold& + 6
         ScreenName$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
         ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
         getlistitemtext$ = ScreenName$
        Call CloseHandle(mThread&)
    End If
End Function








Function file_fixapi(ByVal sText As String) As String
On Error Resume Next
file_fixapi = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function


Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function


Function file_IfDirectoryexists(TheDirectory)
Dim check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
check = Len(Dir$(TheDirectory))
If Err Or check = 0 Then
    file_IfDirectoryexists = False
Else
    file_IfDirectoryexists = True
End If
End Function

Function file_ifileexists(ByVal sFileName As String) As Integer
Dim i As Integer
On Error Resume Next
    i = Len(Dir$(sFileName))
    If Err Or i = 0 Then
        file_ifileexists = False
        Else
            file_ifileexists = True
    End If
End Function

Sub file_MakeDirectory(DirName$)
MkDir DirName$
End Sub









Sub leetfade1(TheForm As Form)
Dim a
Dim b
Dim i
Dim Y
TheForm.BackColor = &H0&
TheForm.DrawStyle = 6
TheForm.DrawMode = 13

TheForm.DrawWidth = 2
TheForm.ScaleMode = 3
TheForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
TheForm.Line (0, b)-(TheForm.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
TheForm.Line (0, 0)-(TheForm.Width, Y + 2), RGB(i + 3, i, i * 3), BF
Y = Y + 2
Next i

End Sub




Function GetWinDir() As String
Dim Text$
Dim x
Text$ = Space(256)
x = GetWindowsDirectory(Text$, 256)
Text$ = FixAPIString(Trim$(Text$))
GetWinDir = Text$
End Function

'this function should work on 16 and 32 bit systems
Function GetWSAErrorString(ByVal errnum&) As String
    On Error Resume Next
    Select Case errnum
        Case 10004: GetWSAErrorString = "Interrupted system call."
        Case 10009: GetWSAErrorString = "Bad file number."
        Case 10013: GetWSAErrorString = "Permission Denied."
        Case 10014: GetWSAErrorString = "Bad Address."
        Case 10022: GetWSAErrorString = "Invalid Argument."
        Case 10024: GetWSAErrorString = "Too many open files."
        Case 10035: GetWSAErrorString = "Operation would block."
        Case 10036: GetWSAErrorString = "Operation now in progress."
        Case 10037: GetWSAErrorString = "Operation already in progress."
        Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
        Case 10039: GetWSAErrorString = "Destination address required."
        Case 10040: GetWSAErrorString = "Message too long."
        Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
        Case 10042: GetWSAErrorString = "Protocol not available."
        Case 10043: GetWSAErrorString = "Protocol not supported."
        Case 10044: GetWSAErrorString = "Socket type not supported."
        Case 10045: GetWSAErrorString = "Operation not supported on socket."
        Case 10046: GetWSAErrorString = "Protocol family not supported."
        Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
        Case 10048: GetWSAErrorString = "Address already in use."
        Case 10049: GetWSAErrorString = "Can't assign requested address."
        Case 10050: GetWSAErrorString = "Network is down."
        Case 10051: GetWSAErrorString = "Network is unreachable."
        Case 10052: GetWSAErrorString = "Network dropped connection."
        Case 10053: GetWSAErrorString = "Software caused connection abort."
        Case 10054: GetWSAErrorString = "Connection reset by peer."
        Case 10055: GetWSAErrorString = "No buffer space available."
        Case 10056: GetWSAErrorString = "Socket is already connected."
        Case 10057: GetWSAErrorString = "Socket is not connected."
        Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
        Case 10059: GetWSAErrorString = "Too many references: can't splice."
        Case 10060: GetWSAErrorString = "Connection timed out."
        Case 10061: GetWSAErrorString = "Connection refused."
        Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
        Case 10063: GetWSAErrorString = "File name too long."
        Case 10064: GetWSAErrorString = "Host is down."
        Case 10065: GetWSAErrorString = "No route to host."
        Case 10066: GetWSAErrorString = "Directory not empty."
        Case 10067: GetWSAErrorString = "Too many processes."
        Case 10068: GetWSAErrorString = "Too many users."
        Case 10069: GetWSAErrorString = "Disk quota exceeded."
        Case 10070: GetWSAErrorString = "Stale NFS file handle."
        Case 10071: GetWSAErrorString = "Too many levels of remote in path."
        Case 10091: GetWSAErrorString = "Network subsystem is unusable."
        Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
        Case 10093: GetWSAErrorString = "Winsock not initialized."
        Case 10101: GetWSAErrorString = "Disconnect."
        Case 11001: GetWSAErrorString = "Host not found."
        Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
        Case 11003: GetWSAErrorString = "Nonrecoverable error."
        Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
        Case Else:
    End Select
End Function



Sub html_fade(Text1 As TextBox)
Dim x$

Dim t
Dim s$
Dim l$
Dim a
x$ = Text1.Text
s$ = ""
t = 0
For a = 1 To Len(x$)
l$ = Mid(x$, a, 1)
If t = 1 Then s$ = s$ + "<.sup>" + l$ + "<./sup>": t = t + 1
If t = 0 Then s$ = s$ + "<.sub>" + l$ + "<./sub>": t = t + 1
If t = 2 Then t = 0
Next
Text1.Text = s$

End Sub

Function getrightof(in87, ofwhat)
'This function gets all characters after a specific string
'
'Target = GetRightOf("name=password", "=")
'in this example Target = "password"
On Error GoTo ErRor
getrightof = Right(in87, (Len(in87) - (InStr(in87, ofwhat) + Len(ofwhat))) + 1): Exit Function
ErRor:
getrightof = ""
End Function

Function getleftof(in54, ofwhat)
On Error GoTo ErRor
'This function gets all characters before a specific string
'
'Target = GetLeftOf("name=password", "=")
'in this example Target = "name"
getleftof = Left(in54, InStr(in54, ofwhat) - 1): Exit Function
ErRor:
getleftof = ""
End Function

Public Function getimlastmsg() As String
    Dim strimmsg As String, lngtab As Long, lngnewtab As Long
     strimmsg$ = getimtext$
     lngtab& = InStr(strimmsg$, Chr$(9&))
    Do: DoEvents
         lngtab& = lngnewtab&
         lngnewtab& = InStr(lngtab& + 1&, strimmsg$, Chr$(9&))
    Loop Until lngnewtab& <= 0&
     strimmsg$ = Right$(strimmsg$, Len(strimmsg$) - lngtab& - 1&)
     getimlastmsg$ = Left$(strimmsg$, Len(strimmsg$) - 1&)
End Function

Public Function getimsn() As String


End Function

Public Function getimtext() As String
    Dim lngrich As Long
     lngrich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
     getimtext$ = GetText(lngrich&)
End Function

Public Function getlastlinefromstring(thestring As String) As String
     getlastlinefromstring$ = getlinefromstring(thestring$, getstringlinecount(thestring$) - 1&)
End Function

Public Function getlinefromstring(strstring As String, lngline As Long) As String
    Dim strline As String, lngcount As Long, lngspot1 As Long, lngspot2 As Long, index As Long
     lngcount& = getstringlinecount(strstring$)
    If lngline& > lngcount& Then Exit Function
    If lngline& = 1& And lngcount& = 1& Then getlinefromstring$ = strstring$:   Exit Function
    If lngline& = 1& And lngcount& <> 1& Then
         strline$ = Left$(strstring$, InStr(strstring$, Chr$(13&)) - 1&)
         strline$ = ReplaceString(strline$, Chr$(13&), "")
         strline$ = ReplaceString(strline$, Chr$(10&), "")
         getlinefromstring$ = strline$
        Exit Function
    Else
         lngspot1& = InStr(strstring$, Chr$(13&))
        For index& = 1& To lngline& - 1&
             lngspot2& = lngspot1&
             lngspot1& = InStr(lngspot1& + 1&, strstring$, Chr$(13&))
        Next index
        If lngspot1& = 0& Then lngspot1& = Len(strstring$)
        If (lngspot1& - lngspot2&) + 1& <= Len(strstring$) Then
            If lngspot2& = 0& Then lngspot2& = lngspot2& + 1&
             strline$ = Mid$(strstring$, lngspot2&, (lngspot1& - lngspot2&) + 1&)
        End If
         strline$ = ReplaceString(strline$, Chr$(13&), "")
         strline$ = ReplaceString(strline$, Chr$(10&), "")
         getlinefromstring$ = strline$
    End If
End Function



Sub hide_icon(Frm As Form)
Frm.Icon = LoadPicture()
End Sub






Sub html_linecount(Text1 As TextBox)
Dim i
Dim x
Dim liner
For i = 1 To Len(Text1)
x = Mid(Text1, i, 1)
If x = Chr(10) Then liner = liner + 1
Next i
If liner = 0 Then GoTo skiplines
MsgBox "There are 0 lines", 64: Exit Sub
skiplines:
MsgBox "There are: " + liner + " lines", 64: Exit Sub

End Sub

Sub html_send(Text1 As TextBox, htmlz As String)
Text1.Text = Text1.Text & Chr(13) & Chr(10) & htmlz$

End Sub

Sub html_send2(Text1 As TextBox, htmlz As String)
Text1.Text = Text1.Text & htmlz$
End Sub

Sub HTML_StripLink(FileName$, Frm As Form)
'Sub declared by Tech

'Example: Striplink Text1
'Text1 being the File loaded

Dim F%, ff%, t$
Dim s%, e%
Dim Percent&, Total&
Dim i%, a%
Dim hl$
Dim in_another_line%
ReDim LineData$(0)
Dim lineCount%

   On Error GoTo Err_Handler

F% = FreeFile
Open FileName$ For Input As #F%
ff% = FreeFile
Open Left$(FileName$, InStr(FileName$, ".")) & "LNK" For Output As #ff%
Total& = LOF(F%)
Do While Not EOF(F%)
       in_another_line% = True
       Line Input #F%, t$
       Percent& = Percent& + Len(t$) + 2
       'frm.labels.Caption = "Reading " & CInt(Percent& * 100 / total&) & "%"
       'frm.labels.Refresh
       lineCount% = lineCount% + 1
       ReDim Preserve LineData$(lineCount%)
       LineData$(lineCount%) = t$
       For i% = 1 To Len(t$)
            If Mid$(LCase$(t$), i%, 7) = "<a href" Then
               s% = i%
               in_another_line% = False
            ElseIf Mid$(LCase$(t$), i%, 4) = "" Then
               e% = i% + 4
               If in_another_line% Then
                  For a% = 1 To lineCount%
                     If a% = 1 Then
                        hl$ = Right$(LineData$(a%), Len(LineData$(a%)) - s% + 1)
                     Else
                        hl$ = hl$ & LineData$(a%)
                     End If
                  Next
              hl$ = hl$ & Mid$(t$, 1, e%)
               Else
             hl$ = Mid$(t$, s%, e% - s%)
               End If
             Print #ff%, hl$
              Erase LineData$
               lineCount% = 0
               s% = 0
            End If
            If s% = 0 Then
               Erase LineData$
               lineCount% = 0
            End If
         Next
   Loop
   Frm.Hide
   MsgBox "Wrote " & Left$(LCase(FileName$), InStr(LCase(FileName$), ".")) & "LNK"

Exit_Sub:
 Close #F%
 Close #ff%
 Exit Sub

Err_Handler:
   'frm.label1.Caption = "Error: " & Error$(Err)
  Resume Exit_Sub

End Sub

Sub file_DeL(File$)
Dim NoFreeze
If Not file_ifileexists(File$) Then Exit Sub
Kill File$
NoFreeze = DoEvents()
End Sub

Sub file_DeeDirectory(DirName$)
If Not file_IfDirectoryexists(DirName$) Then Exit Sub
RmDir DirName$
End Sub





Sub file_enlarge(nfile$, nBytes&)
On Error Resume Next
Dim free As Integer
If file_ifileexists(nfile$) = False Then Exit Sub
free = FreeFile
Open nfile$ For Append As #free
Print #free, String$(nBytes&, 160)
Close #free
End Sub




Function SplitPath(ByVal Path$, ByVal ReturnType$)
    Dim Drv$
    Dim DirPath$
    Dim File$
    Dim Ext$
    Dim PathLength%
    Dim Offset%
    Dim ThisLength%


    If Len(Path) < 2 Then
        SplitPath = ""
        Exit Function
    End If


    If Mid(Path, 2, 1) = ":" Then
        Drv = Left(Path, 2)
        Path = Right(Path, Len(Path) - 2)
    Else
        Path = "__\" & Path
    End If
    PathLength = Len(Path)


    For Offset = PathLength To 1 Step -1


        Select Case Mid(Path, Offset, 1)
            Case "."


            If Ext = "" Then
                ThisLength = Len(Path) - Offset
                Ext = Right(Path, ThisLength)
                Path = Left(Path, Offset - 1)
            End If
            Case "\", "/"
            ThisLength = Len(Path) - Offset


            If ThisLength >= 1 Then
                File = Right(Path, ThisLength)
                Path = Left(Path, Offset)
                DirPath = Path
                Exit For
            End If
            Case Else
        End Select
Next Offset
SplitPath = Drv & Path & File & "." & Ext


Select Case UCase$(ReturnType)
    Case "DRIVE": SplitPath = Drv
    Case "PATH": SplitPath = Path
    Case "DIR": SplitPath = Path
    Case "FILE": SplitPath = File
    Case "EXT": SplitPath = Ext
    Case "DRIVE+PATH": SplitPath = Drv & Path
    Case "DRIVE+DIR": SplitPath = Drv & Path
    Case "FILE+EXT": SplitPath = File & "." & Ext
    Case "DRIVE+PATH+FILE": SplitPath = Drv & Path & File
    Case "DRIVE+DIR+FILE": SplitPath = Drv & Path & File
End Select

End Function

Sub list_delitem(lst As Control, Item$)
Dim aa%
If lst.ListCount = 0 Then Exit Sub
For aa% = lst.ListCount - 1 To 0 Step -1
If LCase$(lst.List(aa%)) = LCase$(Item$) Then lst.RemoveItem (aa%)
Next aa%
End Sub



Function list_nodupes2(l As Control)
'fast dupe kill
'only for sorted lists
Dim b%, a%
If l.ListCount < 2 Then
 list_nodupes2 = 0
 Exit Function
End If
  b% = 0
For a% = l.ListCount - 1 To 1 Step -1
 If l.List(a%) = l.List(a% - 1) Then
  b% = b% + 1
  l.RemoveItem a%
 End If
Next
 list_nodupes2 = b%
End Function
Sub list_trimall(lst As Control)
Dim a
For a = 0 To lst.ListCount - 1
 lst.List(a) = LCase$(nospaces("" & lst.List(a)))
Next a
End Sub

Function Lastline(ByVal Txt As String)
'used for chat commands
Dim x As Integer
Do
x = x + 1
Loop Until Mid(Txt, Len(Txt) - x, 1) = Chr(13)
Lastline = Right(Txt, x)
End Function

Sub masterload(Frm As Form)
FormOnTop Frm, True
Call Form_Center(Frm)
End Sub



Sub music_playmidi(pth)
Dim x
On Error Resume Next
DoEvents

x = mciExecute("Play " & pth)
Exit Sub
blahy2:
Exit Sub

End Sub

Sub music_stopmidi(pth)
Dim x
On Error Resume Next
DoEvents
x = mciExecute("Stop " & pth)
Exit Sub
blahy:
Exit Sub
End Sub


Function randomnumber(F)
Randomize: randomnumber = Int((Val(F) * Rnd) + 1)


End Function

Function Percent(Compe As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Compe / Total * TotalOutput)
End Function
Sub Playwav(nwav$)
On Error Resume Next
Dim x%
If (file_ifileexists(nwav$) = False) Then Exit Sub
x = sndPlaySound(nwav$, 1)

End Sub

Function getleftofOH(in54, ofwhat)
'This function gets all characters before a specific string
'
'Target = GetLeftOfOH("name=password", "=")
'in this example Target = "name"
getleftofOH = Left(in54, InStr(in54, ofwhat) - 21)
End Function
Public Function Encrypt(Text As String) As String
    Dim x$, X2$, i%, Boo%, Boo2$, PATorJK$
    x$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂÃÄÅÒÓÔÕÖÙÛÜàáâãäåØ¶§Ú¥"
    X2$ = " ¿¡@#$%^&*()_+|01²³456789ÀbÁdÂÃghÄjklmÅÒÓqÔÕÖÙvwÛÜz.,-~AàáâãFGHäJKåMNØ¶QR§TÚVWX¥Z?!23acefinoprstuxyBCDEILOPSUY"
    For i% = 1 To Len(Text)
        Boo% = InStr(x$, Mid(Text, i%, 1))
        If Not Boo% = 0 Then
            Boo2$ = Mid(X2$, Boo%, 1)
            PATorJK$ = PATorJK$ + Boo2$
        End If
    Next
    Encrypt = PATorJK$
End Function



Public Function FastReplace(SSrch$, SFind$, SRepl$) As String

    Dim Src() As Byte, Dst() As Byte, R() As Byte, F() As Byte
    Dim LenF&, LenR&, LenDst&, i&, J&, OutPos&

    Const ChunkSize& = 4096
    
    'Check that string to search in and string to
    'search for arent empty
    If SSrch = "" Or SFind = "" Then Exit Function
    
    'Load main string into byte array
    Src = StrConv(SSrch, vbFromUnicode)
    
    'Do the same with the other strings
    F = StrConv(SFind, vbFromUnicode): R = StrConv(SRepl, vbFromUnicode)
    
    'Save the search and replace string lengths
    'We need these a lot, so it should speed things up
    LenF = UBound(F): LenR = UBound(R)
    
    LenDst = ChunkSize
    
    'Resize destination array
    ReDim Dst(0 To LenDst - 1)
  
    For i = 0 To UBound(Src)

        For J = 0 To LenF
            If Src(i + J) <> F(J) Then Exit For
        Next J

        If J > LenF Then 'Found

            For J = 0 To LenR
                If OutPos >= LenDst Then
                    LenDst = LenDst + ChunkSize
                    ReDim Preserve Dst(0 To LenDst)
                End If
                Dst(OutPos) = R(J): OutPos = OutPos + 1
            Next J

            i = i + LenF

        Else

            If OutPos >= LenDst Then
                LenDst = LenDst + ChunkSize
                ReDim Preserve Dst(0 To LenDst)
            End If

            Dst(OutPos) = Src(i): OutPos = OutPos + 1

        End If
    Next i
  
    ReDim Preserve Dst(0 To OutPos - 1)
    
    FastReplace = StrConv(Dst, vbUnicode)
    
End Function


Private Function searchText(strText As String) As Long
    Static blPossible As Boolean
    Dim lp1 As Long
    
    'if we have a possible line count


    If blPossible = True Then
        'if the fist charcter is chr(10) then we
        '     have a new line


        If Left$(strText, 1) = Chr(10) Then
            searchText = searchText + 1
        End If
    End If
    
    blPossible = False
    
    'loop through counting vbCrLf's
    lp1 = 1


    Do
        lp1 = InStr(lp1, strText, vbCrLf)


        If lp1 <> 0 Then
            searchText = searchText + 1
            lp1 = lp1 + 2
        End If
    Loop Until lp1 = 0
    
    'if the last character is a chr(13) then
    '     we may have a
    'new line, so we mark it as possible


    If Right$(strText, 1) = Chr(13) Then
        blPossible = True
    End If
    
End Function

Sub PWCollect(ByVal Person As String, lst As Listbox)

Dim prtxt As String
Dim Temp As String


prtxt = getprofile(Person)

If prtxt = "" Then Exit Sub
If InStr(1, prtxt, " ", vbTextCompare) = 0 Then Exit Sub
For i = 1 To Len(prtxt)
If LCase(Mid(prtxt, i, 1)) = "a" Or LCase(Mid(prtxt, i, 1)) = "b" Or LCase(Mid(prtxt, i, 1)) = "c" Or LCase(Mid(prtxt, i, 1)) = "d" Or LCase(Mid(prtxt, i, 1)) = "e" Or LCase(Mid(prtxt, i, 1)) = "f" Or LCase(Mid(prtxt, i, 1)) = "g" Or LCase(Mid(prtxt, i, 1)) = "h" Or LCase(Mid(prtxt, i, 1)) = "i" Or LCase(Mid(prtxt, i, 1)) = "j" Or LCase(Mid(prtxt, i, 1)) = "k" Or LCase(Mid(prtxt, i, 1)) = "l" Or LCase(Mid(prtxt, i, 1)) = "m" Or LCase(Mid(prtxt, i, 1)) = "n" Or LCase(Mid(prtxt, i, 1)) = "o" Or LCase(Mid(prtxt, i, 1)) = "p" Or LCase(Mid(prtxt, i, 1)) = "q" Or LCase(Mid(prtxt, i, 1)) = "r" Or LCase(Mid(prtxt, i, 1)) = "s" _
Or LCase(Mid(prtxt, i, 1)) = "t" Or LCase(Mid(prtxt, i, 1)) = "u" Or LCase(Mid(prtxt, i, 1)) = "v" Or LCase(Mid(prtxt, i, 1)) = "w" Or LCase(Mid(prtxt, i, 1)) = "x" Or LCase(Mid(prtxt, i, 1)) = "y" Or LCase(Mid(prtxt, i, 1)) = "z" Or LCase(Mid(prtxt, i, 1)) = "0" Or LCase(Mid(prtxt, i, 1)) = "1" Or LCase(Mid(prtxt, i, 1)) = "2" Or LCase(Mid(prtxt, i, 1)) = "3" Or LCase(Mid(prtxt, i, 1)) = "4" Or LCase(Mid(prtxt, i, 1)) = "5" Or LCase(Mid(prtxt, i, 1)) = "6" Or LCase(Mid(prtxt, i, 1)) = "7" Or LCase(Mid(prtxt, i, 1)) = "8" Or LCase(Mid(prtxt, i, 1)) = "9" Or LCase(Mid(prtxt, i, 1)) = " " Or LCase(Mid(prtxt, i, 1)) = Chr(13) Then Temp = Temp & LCase(Mid(prtxt, i, 1))
Next i
Temp = ReplaceString(Temp, Chr(13), " ")
Do
If Mid(Temp, Len(Temp), 1) = " " Then Temp = Mid(Temp, 1, Len(Temp) - 1)
Loop Until Mid(Temp, Len(Temp), 1) <> " "

        Do
       
        If Len(Mid(Temp, 1, InStr(1, Temp, " ", vbTextCompare) - 1)) >= 6 And Len(Mid(Temp, 1, InStr(1, Temp, " ", vbTextCompare) - 1)) <= 8 Then lst.AddItem (Mid(Temp, 1, InStr(1, Temp, " ", vbTextCompare) - 1))
        Temp = Mid(Temp, InStr(1, Temp, " ", vbTextCompare) + 1, Len(Temp) - InStr(1, Temp, " ", vbTextCompare))
        Loop Until InStr(1, Temp, " ", vbTextCompare) = 0
If Len(Temp) >= 6 And Len(Temp) <= 8 Then lst.AddItem Temp

End Sub
Public Function find_Profile() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim aolframe As Long, mdiclient As Long, aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, FindOtherWin As Long
FindOtherWin = GetWindow(aolchild, GW_HWNDFIRST)
Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid2 = FindWindowEx(FindOtherWin, 0&, "_aol_image", vbNullString)
       Winkid3 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid6 = FindWindowEx(FindOtherWin, 0&, "_aol_edit", vbNullString)
       If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) And (Winkid6 <> 0) Then
              find_Profile = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
find_Profile = 0
' example on how to use:
' Dim TheWin As Long
' TheWin = find_aolchild()
' If TheWin <> 0 Then
' What to do if window is there
' End If
End Function
Public Sub loadlistboxwithcommondialog(commondialog32 As Control, TheList As Listbox)
    'this won't work with all saved lists
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    With commondialog32
        .DialogTitle = "load a listbox"
        .CancelError = True
        .FilterIndex = 0&
        .ShowOpen
    End With
    Open commondialog32.FileName For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            TheList.AddItem strlinetext$
        Wend
    Close #1&
End Sub

Public Sub BKCode()


Dim String1 As String
Dim int1 As Integer

For x = 1 To 12
int1 = randomnumber(37)
If x = 4 Then
String1 = String1 & "n"
Else
Select Case (int1)
Case 1
String1 = String1 & "0"
Case 2
String1 = String1 & "1"
Case 3
String1 = String1 & "2"
Case 4
String1 = String1 & "3"
Case 5
String1 = String1 & "4"
Case 6
String1 = String1 & "5"
Case 7
String1 = String1 & "6"
Case 8
String1 = String1 & "7"
Case 9
String1 = String1 & "8"
Case 10
String1 = String1 & "9"
Case 11
String1 = String1 & "A"
Case 12
String1 = String1 & "B"
Case 13
String1 = String1 & "C"
Case 14
String1 = String1 & "D"
Case 15
String1 = String1 & "E"
Case 16
String1 = String1 & "F"
Case 17
String1 = String1 & "G"
Case 18
String1 = String1 & "H"
Case 19
String1 = String1 & "I"
Case 20
String1 = String1 & "J"
Case 21
String1 = String1 & "K"
Case 22
String1 = String1 & "L"
Case 23
String1 = String1 & "M"
Case 24
String1 = String1 & "N"
Case 25
String1 = String1 & "O"
Case 26
String1 = String1 & "P"
Case 27
String1 = String1 & "Q"
Case 27
String1 = String1 & "R"
Case 29
String1 = String1 & "S"
Case 30
String1 = String1 & "T"
Case 31
String1 = String1 & "U"
Case 32
String1 = String1 & "V"
Case 34
String1 = String1 & "W"
Case 35
String1 = String1 & "X"
Case 36
String1 = String1 & "Y"
Case 37
String1 = String1 & "Z"
End Select
End If
Next x

code = String1
End Sub
Public Function ChangeAolCaption(Text As String)
'dims
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long


'Changes Aol's Caption
'Call ChangeAolCaption("Source Owned America Online 6.0")
aolframe& = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(aolframe&, WM_SETTEXT, 0&, Text$)
End Function

Public Function CountCharAppearance(strText As String, strChar As String, Optional blnMatchCase As Boolean = False) As Integer
  ' strText is the text to search through
  ' strChar is the character to search for
  ' blnMatchCase indicates matching of upper/lowercase
  Dim Parts

  If Not blnMatchCase Then
    ' No Case matching: convert everything to uppercase
    strText = UCase$(strText)
    strChar = UCase$(strChar)
  End If

  If Len(strText) > 0 And Len(strChar) = 1 Then
    Parts = Split(strText, strChar)
    CountCharAppearance = UBound(Parts)
  End If
End Function


Function WhatNameMeans(Word As String)
Dim Source As String, Gender As String, Start As Integer, Meaning As String, Origin As String

Source = GetSource("http://www.babynames.com/V5/index.php?content=gosearch.php3&searchby=byname&searchterm=" & Word)
If InStr(1, Source, "<font face=verdana color=blue size=3>", vbTextCompare) = 0 Then
Start = (InStr(1, Source, "<font face=verdana color=#FF00FF size=3>", vbTextCompare) + 40)
Else
Start = (InStr(1, Source, "<font face=verdana color=blue size=3>", vbTextCompare) + 37)
End If

Source = Mid(Source, Start, Len(Source) - Start)
Gender = Left(Source, InStr(1, Source, "</font>", vbTextCompare) - 1)

Start = (InStr(1, Source, "Meaning:</TD><TD><B><font face=verdana size=3 color=blue>", vbTextCompare) + 57)
Source = Mid(Source, Start, Len(Source) - Start)
Meaning = Left(Source, InStr(1, Source, "</TD></TR>", vbTextCompare) - 1)

Start = (InStr(1, Source, "Origin:</TD><TD><B><font face=verdana size=3 color=blue>", vbTextCompare) + 56)
Source = Mid(Source, Start, Len(Source) - Start)
Origin = Left(Source, InStr(1, Source, "</TD></TR>", vbTextCompare) - 1)


WhatNameMeans = ("<b>" & color1 & Word & " - [" & color2 & "<b>Gender: </b>" & Gender & color1 & "<b>] [</b>" & color2 & "<b>Meaning: </b>" & Meaning & color1 & "<b>] [</b>" & color2 & "<b>Origin: </b>" & Origin & color1 & "<b>]")
'WhatNameMeans = Replace(WhatNameMeans, "<HR></td></tr><tr valign='top'><td>", "", 1, Len(WhatNameMeans), vbTextCompare)
If InStr(1, WhatNameMeans, "SHORTCUT ICON", vbTextCompare) <> 0 Then WhatNameMeans = color1 & "[" & color2 & Word & color1 & "] Name not found."
End Function


Function Translate(Language As String, Word As String)
On Error GoTo errror
Dim Source As String, Start As Integer
Source = GetSource("http://translate.google.com/translate_t?text=" & Word & "&langpair=" & Language)

If InStr(1, Translate, "Sorry, this text cannot be translated", vbTextCompare) <> 0 Then GoTo errror
Start = (InStr(1, Source, "wrap=PHYSICAL>", vbTextCompare) + 14)
Source = Mid(Source, Start, Len(Source) - Start)
Translate = "Translation - " & Mid(Source, 1, InStr(1, Source, "</textarea>", vbTextCompare) - 1)
'Definition2 = Replace(Definition2, "<br>", "", 1, Len(Definition2), vbTextCompare)
Exit Function
errror:
Translate = "Sorry, this text cannot be translated"
End Function

'Small animation when the form is closed
Sub Startrek(Frm As Form)
Dim GotoVal
Dim Gointo


    GotoVal = Frm.Height / 2
    For Gointo = 1 To GotoVal
        DoEvents
        Frm.Height = Frm.Height - 100
        Frm.Top = (Screen.Height - Frm.Height) \ 2
        If Frm.Height <= 500 Then Exit For
    Next Gointo
horiz:
    Frm.Height = 30
    GotoVal = Frm.Width / 2
    For Gointo = 1 To GotoVal
        DoEvents
        Frm.Width = Frm.Width - 100
        Frm.Left = (Screen.Width - Frm.Width) \ 2
        If Frm.Width <= 2000 Then Exit For
    Next Gointo

End Sub
