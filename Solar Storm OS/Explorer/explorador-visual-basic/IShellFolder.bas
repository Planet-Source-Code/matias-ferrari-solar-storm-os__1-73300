Attribute VB_Name = "mIShellFolderDefs"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
'
' ==============================================================
' A fairly comprehensive wrapping of the IShellFolder and IEnumIDList interfaces with
' some IUnknown thrown in. Also will do about anything that can be done with a pidl...
'
' Note that "IShellFolder Extended Type Library v1.1" (ISHF_Ex.tlb) included with this
' project, must be present and correctly registered on your system, and referenced by
' this project to allow use of these interfaces.
' ==============================================================
'
' Procedure responsibility of pidl memory, unless specified otherwise:
' - Calling procedures are solely responsible for freeing pidls they create,
'   or receive as a return value from a called procedure.
' - Called procedures always copy pidls received in their params, and
'   *never* free pidl params.

' Global IContextMenu2 interface variable filled in ShowShellContextMenu on
' treeview and listview item right click. Used for menu messages in FrmWndProc.
Public ICtxMenu2 As IContextMenu2

' Retrieves a pointer to the shell's IMalloc interface.
' Returns NOERROR if successful or or E_FAIL otherwise.
Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long

' Retrieves the IShellFolder interface for the desktop folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As IShellFolder) As Long

' Frees memory allocated by the shell
Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

' GetItemID item ID retrieval constants
Public Const GIID_FIRST = 1
Public Const GIID_LAST = -1
'
'  ' ====================================================
'  ' item ID (pidl) structs, just for reference
'  '
'  ' item identifier (relative pidl), allocated by the shell
'  Public Type SHITEMID
'    cb As Integer        ' size of struct, including cb itself
'    abID() As Byte    ' variable length item identifier
'  End Type
'
'  ' fully qualified pidl
'  Public Type ITEMIDLIST
'    mkid As SHITEMID  ' list of item identifers, packed into SHITEMID.abID
'  End Type
'

' ================================================================
' interface procs

' Returns a reference to the IMalloc interface.

Public Function isMalloc() As IMalloc
  Static im As IMalloc
  If (im Is Nothing) Then Call SUCCEEDED(SHGetMalloc(im))
  Set isMalloc = im
End Function

' Returns a reference to the desktop folder's IShellFolder interface.

Public Function isfDesktop() As IShellFolder
  Static isf As IShellFolder
  If (isf Is Nothing) Then Call SUCCEEDED(SHGetDesktopFolder(isf))
  Set isfDesktop = isf
End Function

' Returns the IShellFolder interface ID, {000214E6-0000-0000-C000-000000046}

Public Function IID_IShellFolder() As GUID
  Static iid As GUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214E6, 0, 0)
  IID_IShellFolder = iid
End Function

' Returns the IShellDetails interface ID, {000214EC-0000-0000-C000-000000000046}

Public Function IID_IShellDetails() As GUID
  Static iid As GUID
  If (iid.Data1 = 0) Then Call DEFINE_OLEGUID(iid, &H214EC, 0, 0)
  IID_IShellDetails = iid
End Function

' Rtns the reference count for the specified object...

Public Function ObjRefCount(obj As IShellFolderEx_TLB.IUnknown) As Long
  On Error GoTo IsNothing
  
  ' AddRef the object, will err if Nothing, returning 0
  obj.AddRef
  
  ' Release the object and return the original reference count,
  ' less one reference for the proc local "obj" variable.
  ObjRefCount = obj.Release - 1

IsNothing:
End Function

' ================================================================
' pidl utility procs

' Determines if the specified pidl is the desktop folder's pidl.
' Returns True if the pidl is the desktop's pidl, returns False otherwise.

' The desktop pidl is only a single item ID whose value is 0 (the 2 byte
' zero-terminator, i.e. SHITEMID.abID is empty). Direct descendents of
' the desktop (My Computer, Network Neighborhood) are absolute pidls
' (relative to the desktop) also with a single item ID, but contain values
' (SHITEMID.abID > 0). Drive folders have 2 item IDs, children of drive
' folders have 3 item IDs, etc. All other single item ID pidls are relative to
' the shell folder in which they reside (just like a relative path).

Public Function IsDesktopPIDL(pidl As Long) As Boolean
  
  ' The GetItemIDSize() call will also return 0 if pidl = 0
  If pidl Then IsDesktopPIDL = (GetItemIDSize(pidl) = 0)

End Function

' Returns the size in bytes of the first item ID in a pidl.
' Returns 0 if the pidl is the desktop's pidl or is the last
' item ID in the pidl (the zero terminator), or is invalid.

Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  
  ' If we try to access memory at address 0 (NULL), then it's bye-bye...
  If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2

End Function

' Returns the count of item IDs in a pidl.

Public Function GetItemIDCount(ByVal pidl As Long) As Integer
  Dim nItems As Integer
  
  ' If the size of an item ID is 0, then it's the zero
  ' value terminating item ID at the end of the pidl.
  Do While GetItemIDSize(pidl)
    pidl = GetNextItemID(pidl)
    nItems = nItems + 1
  Loop
  
  GetItemIDCount = nItems

End Function

' Returns a pointer to the next item ID in a pidl.
' Returns 0 if the next item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer   ' SHITEMID.cb, 2 bytes
  
  cb = GetItemIDSize(pidl)
  ' Make sure it's not the zero value terminator.
  If cb Then GetNextItemID = pidl + cb

End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.

Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer
  ' Error handle in case we get a bad pidl and overflow cb.
  ' (most item IDs are roughly 20 bytes in size, and since an item ID represents
  ' a folder, a pidl can never exceed 260 folders, or 5200 bytes).
  On Error GoTo Out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop
    ' Add 2 bytes for the zero terminating item ID
    GetPIDLSize = cb + 2
  End If
  
Out:
End Function

' Copies and returns the specified item ID from a complex pidl
'   pidl -    pointer to an item ID list from which to copy
'   nItem - 1-based position in the pidl of the item ID to copy

' If successful, returns a new item ID (single-element pidl)
' from the specified element positon. Returns 0 on failure.
' If nItem exceeds the number of item IDs in the pidl,
' the last item ID is returned.

' (calling proc is responsible for freeing the new pidl)

Public Function GetItemID(ByVal pidl As Long, ByVal nItem As Integer) As Long
  Dim nCount As Integer
  Dim i As Integer
  Dim cb As Integer
  Dim pidlNew As Long
  
  nCount = GetItemIDCount(pidl)
  If (nItem > nCount) Or (nItem = GIID_LAST) Then nItem = nCount
  
  ' GetNextItemID returns the 2nd item ID
  For i = 1 To nItem - 1: pidl = GetNextItemID(pidl): Next
    
  ' Get the size of the specified item identifier.
  ' If cb = 0 (the zero terminator), the we'll return a desktop pidl, proceed
  cb = GetItemIDSize(pidl)
  
  ' Allocate a new item identifier list.
  pidlNew = isMalloc.Alloc(cb + 2)
  If pidlNew Then
    
    ' Copy the specified item identifier.
    ' and append the zero terminator.
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    MoveMemory ByVal pidlNew + cb, 0, 2
    
    GetItemID = pidlNew
  End If
  
End Function

' Creates a new pidl of the given size

' (calling proc is responsible for freeing the new pidl)

Public Function CreatePIDL(cb As Long) As Long
  Dim pidl As Long
  
  pidl = isMalloc.Alloc(cb)
  If pidl Then
    FillMemory ByVal pidl, cb, 0 ' initialize to zero, set by caller
    CreatePIDL = pidl
  End If

End Function

' Returns a copy of a relative or absolute pidl

' (calling proc is responsible for freeing the new pidl)

Public Function CopyPIDL(pidl As Long) As Long
  Dim cb As Long
  Dim pidlNew As Long
  
  cb = GetPIDLSize(pidl)
  If cb Then
    pidlNew = CreatePIDL(cb)
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    CopyPIDL = pidlNew
  End If

End Function

' Frees the specified pidl and zeros it

Public Sub FreePIDL(pidl As Long)
  On Error GoTo Out
  
  ' Free the pidl and zero it's *value* only
  ' (not what it points to!, i.e. ZeroMemory = FE...)
  If pidl Then isMalloc.Free ByVal pidl

Out:
  If Err And (pidl <> 0) Then
    Call CoTaskMemFree(pidl)
  End If
  
  pidl = 0
  
End Sub

' Copies and returns all but the last item ID from the specified absolute pidl.

'   pidl                - pointer to the pidl from which to copy
'   fFreeOldPidl  - optional flag specifying whether to free and zero the passed pidl

'    ' If successful, returns a new absolute pid (relative to the desktop)
'    ' If either a valid single item ID pidl is passed to this proc (either the
'    ' desktop's pidl or a relative pidl), or an invalid pidl is passed, 0 is returned.

' If successful, returns a new absolute pid (relative to the desktop)
' If either a valid single item ID pidl is passed to this proc (either the
' desktop's pidl or a relative pidl), or an invalid pidl is passed, the
' desktop's pidl is returned.

' (calling proc is responsible for freeing the new pidl)

Public Function GetPIDLParent(pidl As Long, _
                                                  Optional fReturnDesktop As Boolean = False, _
                                                  Optional fFreeOldPidl As Boolean = False) As Long
  Dim nCount As Integer
  Dim pidl1 As Long
  Dim i As Integer
  Dim cb As Integer
  Dim pidlNew As Long
  
  nCount = GetItemIDCount(pidl)
  If (nCount = 0) And (fReturnDesktop = False) Then Exit Function
  
  ' Get the size of all but the pidl's last item ID and zero terminator.
  ' (maintain the value of the original pidl, it's passed ByRef !!)
  pidl1 = pidl
  For i = 1 To nCount - 1
    cb = cb + GetItemIDSize(pidl1)
    pidl1 = GetNextItemID(pidl1)
  Next
  
  ' Allocate a new item ID list with a new terminating 2 bytes.
  pidlNew = isMalloc.Alloc(cb + 2)
  
  ' If the memory was allocated...
  If pidlNew Then
    ' Copy all but the last item ID from the original pidl
    ' to the new pidl and zero the terminating 2 bytes.
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    FillMemory ByVal pidlNew + cb, 2, 0
    
    If fFreeOldPidl Then Call FreePIDL(pidl)
    GetPIDLParent = pidlNew
  
  End If
  
End Function

' Creates a new pidl by prepending pidl2 to pidl1 (i.e pidlNew = pidl1pidl2)

' (calling proc is responsible for freeing the new pidl, the
' two passed pidls are still valid and are not freed unless specified)

Public Function CombinePIDLs(pidl1 As Long, _
                                                  pidl2 As Long, _
                                                  Optional fFreePidl1 As Boolean = False, _
                                                  Optional fFreePidl2 As Boolean = False) As Long
  Dim cb1 As Integer
  Dim cb2 As Integer
  Dim pidlNew As Long

  ' If pidl1 is non-zero...
  If pidl1 Then
    ' Get it's size
    cb1 = GetPIDLSize(pidl1)
    ' If pidl1 is valid (has a size), subtract the size of the zero terminator
    If cb1 Then cb1 = cb1 - 2
  End If
  
  ' If pidl2 is non-zero...
  If pidl2 Then
    ' Get it's size
    cb2 = GetPIDLSize(pidl2)
    ' If pidl2 is valid (has a size), subtract the size of the zero terminator
    If cb2 Then cb2 = cb2 - 2
  End If

  ' Create a new pidl sized to hold both pidl1, pidl2 and the zero terminator
  pidlNew = CreatePIDL(cb1 + cb2 + 2)
  If (pidlNew) Then
    
    ' If pidl1 is valid, put it's id list at the beginning of our new pidl
    If cb1 Then MoveMemory ByVal pidlNew, ByVal pidl1, cb1
    
    ' If pidl2 is valid, prepend it's id list to the end of the new pidl
    If cb2 Then MoveMemory ByVal pidlNew + cb1, ByVal pidl2, cb2
      
    ' Zero the terminating 2 bytes
    FillMemory ByVal pidlNew + cb1 + cb2, 2, 0
      
    ' Finally, free the pidls as specified
    If (pidl1 And fFreePidl1) Then isMalloc.Free ByVal pidl1
    If (pidl2 And fFreePidl2) Then isMalloc.Free ByVal pidl2
    
  End If
  
  CombinePIDLs = pidlNew

End Function

' Returns an absolute pidl's path only (doesn't rtn display names!)

Public Function GetPathFromPIDL(pidl As Long) As String
  Dim sPath As String * MAX_PATH   ' 260
  If SHGetPathFromIDList(pidl, sPath) Then
    GetPathFromPIDL = GetStrFromBufferA(sPath)
  End If
End Function

' ================================================================
' IShellFolder procs

' Returns a shell item's displayname

'   isfParent - item's parent folder IShellFolder
'   pidlRel    - item's pidl, relative to isfParent
'   uFlags    - specifies the type of name to retrieve

Public Function GetFolderDisplayName(isfParent As IShellFolder, _
                                                              pidlRel As Long, _
                                                              uFlags As ESHGNO) As String
  Dim lpStr As STRRET   ' struct filled
  
  If SUCCEEDED(isfParent.GetDisplayNameOf(pidlRel, uFlags, lpStr)) Then
    GetFolderDisplayName = GetStrRet(lpStr, pidlRel)
  End If

End Function

' Returns information from the STRRET struct (identical to the new IE5 StrRetToStr API).

Public Function GetStrRet(lpStr As STRRET, pidlRel As Long) As String
  Dim lpsz As Long         ' string pointer
  Dim uOffset As Long    ' offset to the string pointer
  
  Select Case (lpStr.uType)
  
    ' The 1st UINT (Long) of the array points to a Unicode
    ' str which *should* be allocated & freed.
    Case STRRET_WSTR
      MoveMemory lpsz, lpStr.CStr(0), 4
      GetStrRet = GetStrFromPtrW(lpsz)
      Call CoTaskMemFree(lpsz)
    
    ' The 1st UINT (Long) of the array points to the location
    ' (uOffset bytes) to the ANSI str in the pidl.
    Case STRRET_OFFSET
      MoveMemory uOffset, lpStr.CStr(0), 4
      GetStrRet = GetStrFromPtrA(pidlRel + uOffset)
    
    ' The display name is returned in cStr.
    Case STRRET_CSTR
      GetStrRet = GetStrFromPtrA(VarPtr(lpStr.CStr(0)))
  
  End Select

End Function

' Returns the IShellFolder for the specified relative pidl

'   isfParent - pidl's parent folder IShellFolder
'   pidlRel    - child folder's relative pidl we're returning the IShellFolder of.

' If an error occurs, the desktop's IShellFolder is returned.

Public Function GetIShellFolder(isfParent As IShellFolder, pidlRel As Long) As IShellFolder
  Dim isf As IShellFolder
  On Error GoTo Out
  
  Call isfParent.BindToObject(pidlRel, 0, IID_IShellFolder, isf)

Out:
  If Err Or (isf Is Nothing) Then
    Set GetIShellFolder = isfDesktop
  Else
    Set GetIShellFolder = isf
  End If

End Function

' Returns a reference to the parent IShellFolder of the last item ID in the specified
' fully qualified pidl (identical to the new Win2K SHBindToParent function).

' If pidlFQ is zero, or a relative (single item) pidl, then the desktop's IShellFolder
' is returned. If an unexpected error occurs, the object value Nothing is returned.

Public Function GetIShellFolderParent(ByVal pidlFQ As Long, _
                                                            Optional fRtnDesktop As Boolean = True) As IShellFolder
  Dim pidlParent As Long

  pidlParent = GetPIDLParent(pidlFQ, fRtnDesktop)
  If pidlParent Then
    Set GetIShellFolderParent = GetIShellFolder(isfDesktop, pidlParent)
    isMalloc.Free ByVal pidlParent
  End If

End Function

' Displays the specified items' shell context menu.
'
'    hwndOwner  - window handle that owns context menu and any err msgboxes
'    isfParent       - pointer to the items' parent shell folder
'    cPidls            - count of pidls at, and after, pidlRel
'    pidlRel          - the first item's pidl, relative to isfParent
'    pt                  - location of the context menu, in screen coords
'
' Returns True if a context menu command was selected, False otherwise.

Public Function ShowShellContextMenu(hwndOwner As Long, _
                                                                isfParent As IShellFolder, _
                                                                cPidls As Long, _
                                                                pidlRel As Long, _
                                                                pt As POINTAPI) As Boolean
  Dim IID_IContextMenu As GUID
  Dim IID_IContextMenu2 As GUID
  Dim icm As IContextMenu
  Dim hr As Long   ' HRESULT
  Dim hMenu As Long
  Dim idCmd As Long
  Dim cmi As CMINVOKECOMMANDINFO
    
  ' Fill the IContextMenu interface ID, {000214E4-000-000-C000-000000046}
  Call DEFINE_OLEGUID(IID_IContextMenu, &H214E4, 0, 0)
    
  ' Get a refernce to the item's IContextMenu interface
  hr = isfParent.GetUIObjectOf(hwndOwner, cPidls, pidlRel, IID_IContextMenu, 0, icm)
  If SUCCEEDED(hr) Then
    
    ' Fill the IContextMenu2 interface ID, {000214F4-000-000-C000-000000046}
    ' and get the folder's IContextMenu2. Is needed so the "Send To" and "Open
    ' With" submenus get filled from the HandleMenuMsg call in FrmWndProc.
    Call DEFINE_OLEGUID(IID_IContextMenu2, &H214F4, 0, 0)
    Call icm.QueryInterface(IID_IContextMenu2, ICtxMenu2)
    
    ' Create a new popup menu...
    hMenu = CreatePopupMenu()
    If hMenu Then

      ' Add the item's shell commands to the popup menu.
      If (ICtxMenu2 Is Nothing) = False Then
        hr = ICtxMenu2.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)
      Else
        hr = icm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)
      End If
      If SUCCEEDED(hr) Then
        
        ' Show the item's context menu.
        idCmd = TrackPopupMenu(hMenu, _
                                                    TPM_LEFTBUTTON Or TPM_RIGHTBUTTON Or _
                                                    TPM_LEFTALIGN Or TPM_TOPALIGN Or _
                                                    TPM_HORIZONTAL Or TPM_RETURNCMD, _
                                                    pt.x, pt.y, 0, hwndOwner, 0)
        
        ' If a menu command is selected...
        If idCmd Then
          
          ' Fill the struct with the selected command's information.
          With cmi
            .cbSize = Len(cmi)
            .hwnd = hwndOwner
            .lpVerb = idCmd - 1 ' MAKEINTRESOURCE(idCmd-1);
            .nShow = SW_SHOWNORMAL
          End With

          ' Invoke the shell's context menu command. The call itself does
          ' not err if the pidlRel item is invalid, but depending on the selected
          ' command, Explorer *may* raise an err. We don't need the return
          ' val, which should always be NOERROR anyway...
          If (ICtxMenu2 Is Nothing) = False Then
            Call ICtxMenu2.InvokeCommand(cmi)
          Else
            Call icm.InvokeCommand(cmi)
          End If
          
        End If   ' idCmd
      End If   ' hr >= NOERROR (QueryContextMenu)

      Call DestroyMenu(hMenu)
    
    End If   ' hMenu
  End If   ' hr >= NOERROR (GetUIObjectOf)

  ' Release the folder's IContextMenu2 from the global variable.
  Set ICtxMenu2 = Nothing
  
  ' Returns True if a menu command was selected
  ' (letting us know to explicitly select the right clicked object, if needed)
  ShowShellContextMenu = CBool(idCmd)

End Function
'
'' Returns the list of displaynames for each relative pidl
'' (item ID) in the specified fully qualified pidl (item ID list).
'
'' called from nowhere, a debugging proc.
'
'Public Function GetPIDLNames(pidlFQ As Long) As String
'  Dim nItems As Integer
'  Dim isfParent As IShellFolder
'  Dim i As Integer
'  Dim pidlRel As Long
'  Dim sNames As String
'
'  ' Get the count of item ID's in the item ID list.
'  nItems = GetItemIDCount(pidlFQ)
'  If nItems Then
'
'    ' Start with the desktop's shell folder.
'    Set isfParent = isfDesktop
'
'    ' Walk through the each item ID in the item ID list.
'    For i = 1 To nItems '- 1
'
'      ' Get the current relative pidl (item ID) from the
'      ' fully qualified pidl (item ID list)
'      pidlRel = GetItemID(pidlFQ, i)
'      If pidlRel Then
'
'        ' Append each item ID's displayname to the output string.
'        sNames = sNames & GetFolderDisplayName(isfParent, _
'                                                                              pidlRel, _
'                                                                              SHGDN_INFOLDER) & vbCrLf
'        ' Bind to the current item ID's shell folder,
'        ' setting it as the new parent shell folder
'        If SUCCEEDED(isfParent.BindToObject(pidlRel, 0, IID_IShellFolder, isfParent)) = False Then
'          Exit For
'        End If
'
'        ' Free the relative pidl we just got and
'        ' set it to 0 so we know it's freed.
'        isMalloc.Free ByVal pidlRel
'        pidlRel = 0
'
'      End If   ' pidlRel
'    Next
'
'  End If   ' nItems
'
'  ' If the BindToObject call failed above and we exited
'  ' the For loop without freeing the relative pidl, free it now.
'  If pidlRel Then isMalloc.Free ByVal pidlRel
'
'  ' Return the item ID list's displaynames
'  GetPIDLNames = sNames
'
'End Function
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  