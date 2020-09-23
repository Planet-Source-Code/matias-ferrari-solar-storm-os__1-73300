VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPUUsage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' OS Version API Calls
Private Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformId         As Long
   szCSDVersion         As String * 128
End Type

Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' Windows 9x API Calls
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_DYN_DATA = &H80000006
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS = 0&
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const REG_DWORD As Long = 4

' Windows NT\2000\XP API Calls
Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0

Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type

Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type

Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias  As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' Class Level Variables
Private m_blnIsWinNT As Boolean
Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER
Private hKey As Long
Private dwDataSize As Long
Private dwCpuUsage As Byte
Private dwType As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

 Sub AlwaysOnTop(F As Form, Valor As Boolean)
    
    If Valor Then
        SetWindowPos F.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos F.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

' Returns the current usage of the CPU as a percentage
Public Function CurrentCPUUsage() As Long

   Dim structSysBaseInfo      As SYSTEM_BASIC_INFORMATION
   Dim structSysPerfInfo      As SYSTEM_PERFORMANCE_INFORMATION
   Dim structSysTimeInfo      As SYSTEM_TIME_INFORMATION
   Dim dblIdleTime            As Currency
   Dim dblSystemTime          As Currency
   Dim lngReturn              As Long

   If m_blnIsWinNT = True Then
   
      CurrentCPUUsage = 0
      
      ' get number of processors in the system
      lngReturn = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(structSysBaseInfo), LenB(structSysBaseInfo), 0&)
      
      If lngReturn <> NO_ERROR Then Exit Function
      
      ' get new system time
      lngReturn = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(structSysTimeInfo), LenB(structSysTimeInfo), 0&)
      If lngReturn <> NO_ERROR Then Exit Function

      ' get new CPU's idle time
      lngReturn = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(structSysPerfInfo), LenB(structSysPerfInfo), ByVal 0&)
      If lngReturn <> NO_ERROR Then Exit Function

      ' CurrentValue = NewValue - OldValue
      dblIdleTime = LI2Currency(structSysPerfInfo.liIdleTime) - LI2Currency(liOldIdleTime)
      dblSystemTime = LI2Currency(structSysTimeInfo.liKeSystemTime) - LI2Currency(liOldSystemTime)
      
      ' CurrentCpuIdle = IdleTime / SystemTime
      If dblSystemTime <> 0 Then dblIdleTime = dblIdleTime / dblSystemTime
      
      ' CurrentCpuUsage% = 100 - (CurrentCpuIdle * 100) / NumberOfProcessors
      dblIdleTime = 100 - dblIdleTime * 100 / structSysBaseInfo.bKeNumberProcessors + 0.5
      CurrentCPUUsage = Int(dblIdleTime)
      
      ' store new CPU's idle and system time
      liOldIdleTime = structSysPerfInfo.liIdleTime
      liOldSystemTime = structSysTimeInfo.liKeSystemTime
   
   Else
   
      dwDataSize = 4 'size of Long
      
      'Query the counter
      RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
      CurrentCPUUsage = CLng(dwCpuUsage)
   
   End If

End Function

Private Sub Class_Initialize()

   Dim structOSVersionInfo    As OSVERSIONINFO
   Dim structSysTimeInfo      As SYSTEM_TIME_INFORMATION
   Dim structSysPerfInfo      As SYSTEM_PERFORMANCE_INFORMATION
   Dim lngReturn              As Long
   
   structOSVersionInfo.dwOSVersionInfoSize = Len(structOSVersionInfo)
   GetVersionEx structOSVersionInfo
   m_blnIsWinNT = (structOSVersionInfo.dwPlatformId = 2)
   
   ' Check which version of windows (9x or NT\2K\XP)
   If m_blnIsWinNT = True Then
   
      'get new system time
      lngReturn = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(structSysTimeInfo), LenB(structSysTimeInfo), 0&)
      If lngReturn <> NO_ERROR Then Exit Sub

      'get new CPU's idle time
      lngReturn = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(structSysPerfInfo), LenB(structSysPerfInfo), ByVal 0&)
      If lngReturn <> NO_ERROR Then Exit Sub

      'store new CPU's idle and system time
      liOldIdleTime = structSysPerfInfo.liIdleTime
      liOldSystemTime = structSysTimeInfo.liKeSystemTime
   
   Else
    
      ' start the counter by reading the value of the 'StartStat' key
      If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StartStat", 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Sub

      dwDataSize = 4 'Length of Long
      RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
      RegCloseKey hKey
      
      ' get current counter's value
      If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Sub
    
   End If

End Sub

Private Sub Class_Terminate()
    
   If m_blnIsWinNT = False Then
       
      RegCloseKey hKey
      
      'stopping the counter
      If RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StopStat", 0, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Sub
      dwDataSize = 4 'length of Long
      RegQueryValueEx hKey, "KERNEL\CPUUsage", ByVal 0&, dwType, dwCpuUsage, dwDataSize
      RegCloseKey hKey

   End If

End Sub

Private Function LI2Currency(liInput As LARGE_INTEGER) As Currency
   
   CopyMemory LI2Currency, liInput, LenB(liInput)
   
End Function
