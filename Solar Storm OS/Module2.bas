Attribute VB_Name = "Module2"


Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String)
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)
Global myVer As String
Global status$
Global UpdateTime As Integer

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#If Win32 Then


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
#Else


Public Declare Function ShellExecute Lib "shell.dll" (ByVal hwnd As Integer, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
#End If
Public Const SW_SHOWNORMAL = 1

Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
On Local Error GoTo 100
    
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    myfile$ = DestDIR + "\" + RealFile$
    Open myfile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    
    Exit Function

100 X = MsgBox("An error has occured in the file transfer or write.  Please try again later.", vbInformation)
    GetInternetFile = False
    Resume 105
105 End Function
