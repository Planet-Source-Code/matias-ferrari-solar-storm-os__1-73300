Attribute VB_Name = "mWindowDefs"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font

' ============================================================================
' common control definitions

Public Const NM_FIRST = -0&   ' (0U-  0U)       ' // generic to all controls
Public Const NM_DBLCLK = (NM_FIRST - 3)
Public Const NM_RETURN = (NM_FIRST - 4)
Public Const NM_RCLICK = (NM_FIRST - 5)

' The NMHDR structure contains information about a notification message. The pointer
' to this structure is specified as the lParam member of the WM_NOTIFY message.
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long        ' Identifier of control sending message
  code  As Long          ' Specifies the notification code
End Type

' Callback constants

' TV/LV_ITEM.pszText
Public Const LPSTR_TEXTCALLBACK = (-1)

' TVITEM.iImage/iSelectedImage, LVITEM.iImage
Public Const I_IMAGECALLBACK = (-1)

' OCM_NOTIFY is WM_NOTIFY reflected to a C++ created ActiveX control.
' http://msdn.microsoft.com/library/devprods/vs6/visualc/vccore/_core_activex_controls.3a_.subclassing_a_windows_control.htm
Public Const WM_NOTIFY = &H4E
Public Const OCM__BASE = (WM_USER + &H1C00)
Public Const OCM_NOTIFY = (OCM__BASE + WM_NOTIFY)

' ============================================================================
' window messages

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long   ' <---

Public Const WM_DESTROY = &H2
Public Const WM_CANCELMODE = &H1F
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_INITMENUPOPUP = &H117

' ============================================================
' imagelist definitions

Declare Function FileIconInit Lib "shell32.dll" Alias "#660" (ByVal cmd As Boolean) As Boolean

' transparent color (the imagelist will use each icon's mask)
Public Const CLR_NONE = &HFFFFFFFF
Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal himl As Long, ByVal clrBk As Long) As Long
Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long

' ============================================================================
' general window definitions

Public Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum

Public Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Public Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'Declare Function GetFocus Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As CBoolean) As CBoolean

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' ============================================================================
' string/kernel32 definitions

' Converts a Unicode str to a ANSII str.
' Specify -1 for cchWideChar and 0 for cchMultiByte to rtn str len.
Declare Function WideCharToMultiByte Lib "kernel32" _
                            (ByVal CodePage As Long, _
                            ByVal dwFlags As Long, _
                            lpWideCharStr As Any, _
                            ByVal cchWideChar As Long, _
                            lpMultiByteStr As Any, _
                            ByVal cchMultiByte As Long, _
                            ByVal lpDefaultChar As String, _
                            ByVal lpUsedDefaultChar As Long) As Long
' CodePage
Public Const CP_ACP = 0        ' ANSI code page
Public Const CP_OEMCP = 1   ' OEM code page

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, _
                            lpSource As Any, _
                            ByVal dwMessageId As Long, _
                            ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, _
                            ByVal nSize As Long, _
                            Arguments As Long) As Long
' dwFlags
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

' dwLanguageId
Public Const LANG_USER_DEFAULT = &H400&

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

' Loads a string resource from the executable file associated with a specified module
Declare Function LoadString Lib "user32" Alias "LoadStringA" _
                            (ByVal hInstance As Long, _
                            ByVal uID As Long, _
                            ByVal lpBuffer As String, _
                            ByVal nBufferMax As Long) As Long

Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Declare Function lstrcmpiA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function lstrcpyW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long
Declare Function lstrcmpiW Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (pDest As Any, ByVal dwLength As Long, ByVal bFill As Byte)

' =================================================================
' FindFirstFile definitions

'Public Const MAX_PATH = 260

Public Type FILETIME   ' ft
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA   ' wfd
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternateFileName As String * 14
End Type

' nFileSizeHigh: Specifies the high-order DWORD value of the file size, in bytes.
' This value is zero unless the file size is greater than MAXDWORD. The size of
' the file is equal to (nFileSizeHigh * MAXDWORD) + nFileSizeLow.
Public Const MAXDWORD = (2 ^ 32) - 1   ' 0xFFFFFFFF

Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Boolean

'FindFirstFile error rtn value
Public Const INVALID_HANDLE_VALUE = -1

' =================================================================
' file/time definitions

Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long

' Local IDs
Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_USER_DEFAULT = &H400

' Date Flag for GetDateFormat, Time Flag for GetTimeFormat
Public Const LOCALE_NOUSEROVERRIDE = &H80000000    ' do not use user overrides

' Date Flags for GetDateFormat
Public Const DATE_SHORTDATE = &H1                  ' use short date picture
Public Const DATE_LONGDATE = &H2                     ' use long date picture
Public Const DATE_USE_ALT_CALENDAR = &H4   ' use alternate calendar (if any)

' Time Flags for GetTimeFormat
Public Const TIME_NOMINUTESORSECONDS = &H1  ' do not use minutes or seconds
Public Const TIME_NOSECONDS = &H2                        ' do not use seconds
Public Const TIME_NOTIMEMARKER = &H4                 ' do not use time marker, i.e AM/PM
Public Const TIME_FORCE24HOURFORMAT = &H8     ' always use 24 hour format

' ============================================================================
' menu definitions

Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As TPM_wFlags, _
                                                                      ByVal x As Long, ByVal y As Long, _
                                                                      ByVal nReserved As Long, ByVal hwnd As Long, _
                                                                      lprc As Any) As Long
Public Enum TPM_wFlags
  TPM_LEFTBUTTON = &H0
  TPM_RIGHTBUTTON = &H2
  TPM_LEFTALIGN = &H0
  TPM_CENTERALIGN = &H4
  TPM_RIGHTALIGN = &H8
  TPM_TOPALIGN = &H0
  TPM_VCENTERALIGN = &H10
  TPM_BOTTOMALIGN = &H20

  TPM_HORIZONTAL = &H0         ' Horz alignment matters more
  TPM_VERTICAL = &H40            ' Vert alignment matters more
  TPM_NONOTIFY = &H80           ' Don't send any notification msgs
  TPM_RETURNCMD = &H100
End Enum

Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function ShellExecuteEx Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFO) As Long

Public Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hwnd As Long
  lpVerb As Long   ' String
  lpFile As Long   ' String
  lpParameters As Long   ' String
  lpDirectory As Long   ' String
  nShow As Long
  hInstApp As Long
  '  Optional fields
  lpIDList As Long
  lpClass As Long   ' String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

' SHELLEXECUTEINFO fMask
Public Const SEE_MASK_INVOKEIDLIST = &HC

' SHELLEXECUTEINFO nShow
Public Const SW_SHOWNORMAL = 1
'

' Returns the low 16-bit integer from a 32-bit long integer

Public Function LOWORD(dwValue As Long) As Integer
  MoveMemory LOWORD, dwValue, 2
End Function

' Returns the low 16-bit integer from a 32-bit long integer

Public Function HIWORD(dwValue As Long) As Integer
  MoveMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function

' Returns the system-defined description of an API error code

Public Function GetAPIErrStr(dwErrCode As Long) As String
  Dim sErrDesc As String * 256   ' max string resource len
  Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS Or _
                                 FORMAT_MESSAGE_MAX_WIDTH_MASK, _
                                 ByVal 0&, dwErrCode, LANG_USER_DEFAULT, _
                                 ByVal sErrDesc, 256, 0)
  GetAPIErrStr = GetStrFromBufferA(sErrDesc)
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.

Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Returns an ANSII string from a pointer to an ANSII string.

Public Function GetStrFromPtrA(lpszA As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal sRtn, ByVal lpszA)
  GetStrFromPtrA = sRtn
End Function

' Returns an ANSI string from a pointer to a Unicode string.

Public Function GetStrFromPtrW(lpszW As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
'  sRtn = String$(WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, 0, 0, 0, 0), 0)
  Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
  GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function

' Fills a GUID

Public Sub DEFINE_GUID(name As GUID, l As Long, w1 As Integer, w2 As Integer, _
                                          b0 As Byte, b1 As Byte, b2 As Byte, b3 As Byte, _
                                          b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
  With name
    .Data1 = l
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = b0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = b3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
End Sub

' Fills an OLE GUID, the Data4 member always is "C000-000000046"

Public Sub DEFINE_OLEGUID(name As GUID, l As Long, w1 As Integer, w2 As Integer)
  DEFINE_GUID name, l, w1, w2, &HC0, 0, 0, 0, 0, 0, 0, &H46
End Sub

' Provides a generic test for success on any status value.
' Non-negative numbers indicate success.

' If we incur any error situation from any API or interface member
' function's call to this proc, we'll let the user know that sometime's
' not right. What happens when execution continues after the error
' is indeternimate. and could possibly lead to a GPF...

Public Function SUCCEEDED(hr As Long) As Boolean   ' hr = HRESULT
  If (hr >= S_OK) Then
    SUCCEEDED = True
  Else
    If IsIDE Then
      If (MsgBox("Error: &H" & Hex(hr) & ", " & GetAPIErrStr(hr) & vbCrLf & vbCrLf & _
                        "View offending code?", vbExclamation Or vbYesNo) = vbYes) Then Stop
      ' hit Ctrl+L to view the call stack...
    Else
      MsgBox "Error: &H" & Hex(hr) & ", " & GetAPIErrStr(hr), vbExclamation
    End If
  End If
End Function

Public Function IsIDE() As Boolean
  On Error GoTo Out
  Debug.Print 1 / 0
Out:
  IsIDE = Err
End Function

' A dummy procedure that receives and returns the result
' of the AddressOf operator

Public Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function

' Returns the top level parent window from the specified window handle.

Public Function GetTopLevelParent(hwnd As Long) As Long
  Dim hwndParent As Long
  Dim hwndTmp As Long
  
  hwndParent = hwnd
  Do
    hwndTmp = GetParent(hwndParent)
    If hwndTmp Then hwndParent = hwndTmp
  Loop While hwndTmp

  GetTopLevelParent = hwndParent

End Function

' rtns date/time string as "m/d/yy h:m AM/PM"

Public Static Function GetFileDateTimeStr(ftFile As FILETIME) As String
  Dim ftLocal As FILETIME
  Dim st As SYSTEMTIME

  Call FileTimeToLocalFileTime(ftFile, ftLocal)
  Call FileTimeToSystemTime(ftLocal, st)
  GetFileDateTimeStr = GetFileDateStr(st) & " " & GetFileTimeStr(st)

End Function

Public Static Function GetFileDateStr(st As SYSTEMTIME) As String
  Dim sDate As String * 32
  Dim wLen As Integer
  
  wLen = GetDateFormat(LOCALE_USER_DEFAULT, _
                                        LOCALE_NOUSEROVERRIDE Or DATE_SHORTDATE, _
                                        st, vbNullString, sDate, 64)
  
  If wLen Then GetFileDateStr = Left$(sDate, wLen - 1)
  
End Function

Public Static Function GetFileTimeStr(st As SYSTEMTIME) As String
  Dim sTime As String * 32
  Dim wLen As Integer
  
  wLen = GetTimeFormat(LOCALE_USER_DEFAULT, _
                                        LOCALE_NOUSEROVERRIDE Or TIME_NOSECONDS, _
                                        st, vbNullString, sTime, 64)
  
  If wLen Then GetFileTimeStr = Left$(sTime, wLen - 1)
  
End Function

' Returns the string resource contained within the specifed module
' from the specified string resource ID.

Public Function GetResourceString(sModule As String, idString As Long) As String
  Dim hModule As Long
  Dim sBuf As String * MAX_PATH
  Dim nChars As Long
  
  hModule = LoadLibrary(sModule)
  If hModule Then
    nChars = LoadString(hModule, idString, sBuf, MAX_PATH)
    If nChars Then GetResourceString = Left$(sBuf, nChars)
    Call FreeLibrary(hModule)
  End If
  
End Function
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             