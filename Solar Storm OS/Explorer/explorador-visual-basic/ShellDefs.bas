Attribute VB_Name = "mShellDefs"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' General purpose shell definitons found in Shlobj.h
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
' ==============================================================

' shell32.dll string resource IDs, common to all file versions
Public Const IDS_SHELL32_EXPLORE = 8502     ' < Win2K: "&Explore", > Win2K: "E&xplore"
Public Const IDS_SHELL32_NAME = 8976           ' "Name"
Public Const IDS_SHELL32_SIZE = 8978             ' "Size"
Public Const IDS_SHELL32_TYPE = 8979            ' "Type"
Public Const IDS_SHELL32_MODIFIED = 8980    ' "Modified"

Public Const S_OK = 0           ' indicates success
Public Const S_FALSE = 1&   ' special HRESULT value

' Defined as an HRESULT that corresponds to S_OK.
Public Const NOERROR = 0

' Converts an item identifier list to a file system path.
' Returns TRUE if successful or FALSE if an error occurs, for example,
' if the location specified by the pidl parameter is not part of the file system.
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                              (ByVal pidl As Long, _
                              ByVal pszPath As Any) As Long


' ==============================================================
' SHBrowseForFolder

Public Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As String ' Return display name of item selected.
  lpszTitle As String              ' text to go in the banner over the tree.
  ulFlags As Long                 ' Flags that control the return stuff
  lpfn As Long
  lParam As Long      ' extra info that's passed back in callbacks
  iImage As Long      ' output var: where to return the Image index.
End Type

Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                              (lpBrowseInfo As BROWSEINFO) As Long

' typedef int (CALLBACK* BFFCALLBACK)(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData) as long

Public Enum BF_Flags
' Browsing for directory.
  BIF_RETURNONLYFSDIRS = &H1      ' For finding a folder to start document searching
  BIF_DONTGOBELOWDOMAIN = &H2     ' For starting the Find Computer
  ' Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
  ' this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
  ' rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
  ' all three lines of text.
  BIF_STATUSTEXT = &H4
  BIF_RETURNFSANCESTORS = &H8

#If (WIN32_IE >= &H400) Then
  BIF_EDITBOX = &H10               ' Add an editbox to the dialog.  Always on with BIF_USENEWUI
  BIF_VALIDATE = &H20              ' insist on valid result (or CANCEL)
  BIF_USENEWUI = &H40              ' Use the new dialog layout with the ability to resize.
#End If  ' // WIN32_IE >= &H400

  BIF_BROWSEFORCOMPUTER = &H1000  ' Browsing for Computers.
  BIF_BROWSEFORPRINTER = &H2000   ' Browsing for Printers
  BIF_BROWSEINCLUDEFILES = &H4000 ' Browsing for Everything
End Enum

' message from browser
Public Enum BFFM_FromDlg
  BFFM_INITIALIZED = 1
  BFFM_SELCHANGED = 2

#If (WIN32_IE >= &H400) Then
' If the user types an invalid name into the edit box, the browse dialog will call the
' application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message.
' This flag is ignored if BIF_EDITBOX is not specified.
  BFFM_VALIDATEFAILEDA = 3     ' lParam:szPath ret:1(cont),0(EndDialog)
  BFFM_VALIDATEFAILEDW = 4     ' lParam:wzPath ret:1(cont),0(EndDialog)
#End If  ' // WIN32_IE >= &H400
End Enum

' defined in mWindowDefs
Public Const WM_USER = &H400

' messages to browser
Public Enum BFFM_ToDlg
  BFFM_SETSTATUSTEXTA = (WM_USER + 100)
  BFFM_ENABLEOK = (WM_USER + 101)
  BFFM_SETSELECTIONA = (WM_USER + 102)
  BFFM_SETSELECTIONW = (WM_USER + 103)
  BFFM_SETSTATUSTEXTW = (WM_USER + 104)
End Enum

' ==============================================================
' SHGetFileInfo

Public Const MAX_PATH = 260

Public Type SHFILEINFO   ' shfi
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

' Retrieves information about an object in the file system, such as a file,
' a folder, a directory, or a drive root.
Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                              (ByVal pszPath As Any, _
                              ByVal dwFileAttributes As Long, _
                              psfi As SHFILEINFO, _
                              ByVal cbFileInfo As Long, _
                              ByVal uFlags As Long) As Long

Public Enum SHGFI_flags
  SHGFI_LARGEICON = &H0            ' sfi.hIcon is large icon
  SHGFI_SMALLICON = &H1            ' sfi.hIcon is small icon
  SHGFI_OPENICON = &H2              ' sfi.hIcon is open icon
  SHGFI_SHELLICONSIZE = &H4      ' sfi.hIcon is shell size (not system size), rtns BOOL
  SHGFI_PIDL = &H8                        ' pszPath is pidl, rtns BOOL
  ' Indicates that the function should not attempt to access the file specified by pszPath.
  ' Rather, it should act as if the file specified by pszPath exists with the file attributes
  ' passed in dwFileAttributes. This flag cannot be combined with the SHGFI_ATTRIBUTES,
  ' SHGFI_EXETYPE, or SHGFI_PIDL flags <---- !!!
  SHGFI_USEFILEATTRIBUTES = &H10   ' pretend pszPath exists, rtns BOOL
  SHGFI_ICON = &H100                    ' fills sfi.hIcon, rtns BOOL, use DestroyIcon
  SHGFI_DISPLAYNAME = &H200    ' isf.szDisplayName is filled (SHGDN_NORMAL), rtns BOOL
  SHGFI_TYPENAME = &H400          ' isf.szTypeName is filled, rtns BOOL
  SHGFI_ATTRIBUTES = &H800         ' rtns IShellFolder::GetAttributesOf  SFGAO_* flags
  SHGFI_ICONLOCATION = &H1000   ' fills sfi.szDisplayName with filename
                                                        ' containing the icon, rtns BOOL
  SHGFI_EXETYPE = &H2000            ' rtns two ASCII chars of exe type
  SHGFI_SYSICONINDEX = &H4000   ' sfi.iIcon is sys il icon index, rtns hImagelist
  SHGFI_LINKOVERLAY = &H8000    ' add shortcut overlay to sfi.hIcon
  SHGFI_SELECTED = &H10000        ' sfi.hIcon is selected icon
  SHGFI_ATTR_SPECIFIED = &H20000    ' get only attributes specified in sfi.dwAttributes
End Enum
'

' ==============================================================
' SHGetFileInfo calls

' If successful returns the specified file's typename,
' returns an empty string otherwise.
'   pidl  - file's absolute pidl

Public Function GetFileTypeNamePIDL(pidl As Long) As String
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_TYPENAME) Then
    GetFileTypeNamePIDL = GetStrFromBufferA(sfi.szTypeName)
  End If
End Function

' Returns a file's small or large icon index within the system imagelist.
'   pidl       - file's absolute pidl
'   uType  - either SHGFI_SMALLICON or SHGFI_LARGEICON, and SHGFI_OPENICON

Public Function GetFileIconIndexPIDL(pidl As Long, uType As Long) As Long
  Dim sfi As SHFILEINFO
  If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_SYSICONINDEX Or uType) Then
    GetFileIconIndexPIDL = sfi.iIcon
  End If
End Function

' Returns the handle of the small or large icon system imagelist.
'   uSize - either SHGFI_SMALLICON or SHGFI_LARGEICON

Public Function GetSystemImagelist(uSize As Long) As Long
  Dim sfi As SHFILEINFO
  ' Any valid file system path can be used to retrieve system image list handles.
  GetSystemImagelist = SHGetFileInfo("C:\", 0, sfi, Len(sfi), SHGFI_SYSICONINDEX Or uSize)
End Function

' ==============================================================
' SHBrowseForFolder

Public Function BrowseDialog(hwnd As Long, _
                                                sPrompt As String, _
                                                ulFlags As BF_Flags, _
                                                Optional pidlRoot As Long = 0, _
                                                Optional pidlPreSel As Long = 0) As Long
  Dim bi As BROWSEINFO
  
  With bi
    .hwndOwner = hwnd
    .pidlRoot = pidlRoot
    .lpszTitle = sPrompt
    .ulFlags = ulFlags
    .lParam = pidlPreSel
    .lpfn = FARPROC(AddressOf BrowseCallbackProc)
  End With
  
  BrowseDialog = SHBrowseForFolder(bi)
  
End Function

Public Function BrowseCallbackProc(ByVal hwnd As Long, _
                                                            ByVal uMsg As Long, _
                                                            ByVal lParam As Long, _
                                                            ByVal lpData As Long) As Long
'  Dim sPath As String * MAX_PATH
  
  Select Case uMsg
    
    Case BFFM_INITIALIZED
      ' Set the dialog's pre-selected folder from the pidl we set
      ' bi.lParam to above (passed in the lpData param).
      Call SendMessage(hwnd, BFFM_SETSELECTIONA, ByVal CFalse, ByVal lpData)
      
'    Case BFFM_SELCHANGED
'      If SHGetPathFromIDList(lParam, sPath) Then
'        ' Return the path
'        Debug.Print Left$(sPath, InStr(sPath, vbNullChar) - 1)
'      End If
    
  End Select

End Function
