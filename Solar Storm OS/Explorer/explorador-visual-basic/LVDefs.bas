Attribute VB_Name = "mListviewDefs"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' Procedure responsibility of pidl memory, unless specified otherwise:
' - Calling procedures are solely responsible for freeing pidls they create,
'   or receive as a return value from a called procedure.
' - Called procedures always copy pidls received in their params, and
'   *never* free pidl params.

#Const WIN32_IE = &H300

' Holds the relative pidls of the shell items represented by each
' ListItem in the ListView. The string value of each ListItem's
' LVITEM lParam member is its respective collection key.
Private m_colPidlRels As New Collection

' Reference to the selected TreeView folder's IShellFolder,
' used in ListViewCompareItems when soting the relative pidls
Private m_isfParentFolder As IShellFolder

' Flag passed to and used in ListViewCompareProc
' indicating a descending sort (lvwDescending)
Public Const SORT_DESCENDING = &H80000000

' ================================================================
' listview definitions

' style
Public Const LVS_SHAREIMAGELISTS = &H40

' value returned by many listview messages indicating
' the index of no listview item (user defined)
Public Const LVI_NOITEM = &HFFFFFFFF

' messages
Public Const LVM_FIRST = &H1000
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_GETITEM = (LVM_FIRST + 5)
Public Const LVM_SETITEM = (LVM_FIRST + 6)
Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
Public Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Public Const LVM_HITTEST = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Public Const LVM_SORTITEMS = (LVM_FIRST + 48)

' LVM_GET/SETIMAGELIST wParam
Public Const LVSIL_NORMAL = 0
Public Const LVSIL_SMALL = 1

' LVM_GET/SETITEM lParam
Public Type LVITEM   ' was LV_ITEM
  mask As Long
  iItem As Long
  iSubItem As Long
  state As Long
  stateMask As Long
  pszText As Long  ' if String, must be pre-allocated
  cchTextMax As Long
  iImage As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iIndent As Long
#End If
End Type

' LVITEM mask
Public Const LVIF_IMAGE = &H2
Public Const LVIF_PARAM = &H4
Public Const LVIF_STATE = &H8

' LVITEM state, stateMask, LVM_SETCALLBACKMASK wParam
Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const LVIS_OVERLAYMASK = &HF00

' LVM_GETNEXTITEM LOWORD(lParam)
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2

' LVM_GETITEMRECT rc.Left (lParam)
Public Const LVIR_SELECTBOUNDS = 3

' LVM_HITTEST lParam
Public Type LVHITTESTINFO   ' was LV_HITTESTINFO
  pt As POINTAPI
  flags As LVHT_flags
  iItem As Long
#If (WIN32_IE >= &H300) Then
  iSubItem As Long    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
#End If
End Type
 
Public Enum LVHT_flags
'  LVHT_NOWHERE = &H1   ' in LV client area, but not over item
  LVHT_ONITEMICON = &H2
  LVHT_ONITEMLABEL = &H4
  LVHT_ONITEMSTATEICON = &H8
  LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)
 
'  ' outside the LV's client area
'  LVHT_ABOVE = &H8
'  LVHT_BELOW = &H10
'  LVHT_TORIGHT = &H20
'  LVHT_TOLEFT = &H40
End Enum
'
'' notifications
'Public Const LVN_FIRST = -100&  ' (0U-100U)
'Public Const LVN_DELETEITEM = (LVN_FIRST - 3)   ' lParam = NMLISTVIEW
'Public Const LVN_GETDISPINFO = (LVN_FIRST - 50)
'
'' lParam for most listview notification messages
'Public Type NMLISTVIEW   ' was NM_LISTVIEW
'  hdr As NMHDR
'  iItem As Long
'  iSubItem As Long
'  uNewState As Long
'  uOldState As Long
'  uChanged As Long
'  ptAction As POINTAPI
'  lParam As Long
'End Type
'
'' LVN_GETDISPINFO lParam
'Public Type NMLVDISPINFO   ' was LV_DISPINFO
'  hdr As NMHDR
'  item As LVITEM
'End Type
'

' ================================================================
' user-defined listview procedures

' Initializes and fills the specified ListView with the contents of the specified parent folder

'   objLV             - ListView object reference
'   pidlfqParent   - parent folder's fully qualified pidl, relative to the desktop IShellFolder
'   pidlrelParent   - parent folder's pidl, relative to its *parent* IShellFolder

' called only from DoTVNotify/TVN_SELCHANGED

Public Sub FillListView(objLV As ListView, _
                                    pidlfqParent As Long, _
                                    pidlrelParent As Long)
  Dim hwndOwner As Long
  Dim isfParent As IShellFolder
  Dim isdParent As IShellDetails
  Dim shci As SHColInfo   ' (is actually defined in the docs as the SHELLDETAILS struct)
  Dim i As Integer
  Dim ulAttrs As ESFGAO
  Dim ieidl As IEnumIDList
  Dim pidlrelChild As Long
  Dim pidlfqChild As Long
  
  Screen.MousePointer = vbHourglass
  hwndOwner = GetTopLevelParent(objLV.hWnd)

  ' Get a reference to the parent folder's IShellFolder
  Set isfParent = GetIShellFolder(isfDesktop, pidlfqParent)
  
  ' Get the parent folder's SFGAO_FILESYSTEM attribute bit, so we know
  ' what columns to use if the parent folder does not support IShellDetails.
  ulAttrs = SFGAO_FILESYSTEM
  Call GetIShellFolderParent(pidlfqParent).GetAttributesOf(1, pidlrelParent, ulAttrs)
  
  ' First empty the ListView (all relative pidls that we stored from the EnumObjects
  ' enumeration below MUST be freed, or we'll leak memory big time)
  Call ClearListView(objLV)
  
  ' ======================================================
  ' Add column headers specific to the selected folder.
  
  With objLV
    ' Get a reference to the parent folder's IShellDetails
    If (isfParent.CreateViewObject(hwndOwner, IID_IShellDetails, isdParent) = NOERROR) Then
      
      ' Add each column header per IShellDetails info.
      Do While (isdParent.GetDetailsOf(0, i, shci) = NOERROR)
        Call .ColumnHeaders.Add(, , GetStrRet(shci.Text, pidlrelParent), , shci.justify)
        ' SHColInfo.Width (or SHELLDETAILS.cxChar, didn't have the docs
        ' when the IShellDetails interface was added to ISHF_Ex.tlb and
        ' was using undocumented defs) is, according to the docs "Number
        ' of average-sized characters in the header". But multiplying this
        ' value x 6 seems to set a fairly acceptable pixel column width...(?).
'Debug.Print "shci.Width: " & shci.Width & " (" & shci.Width * 6 & ", " & Form1.TextWidth(.ColumnHeaders(i + 1)) & ")"
        Call ListView_SetColumnWidth(.hWnd, (i), shci.Width * 6)
        i = i + 1
      Loop
    
    Else
      ' The parent folder does not implement IShellDetails, explicitly set the
      ' columns... (folders are not required to implement IShellDetails, but
      ' most do, including all non-namespace extension file system folders)
      
      ' Always add the Name column.
      Call .ColumnHeaders.Add(, , GetResourceString("shell32.dll", IDS_SHELL32_NAME), 120, lvwColumnLeft)
      
      If (ulAttrs And SFGAO_FILESYSTEM) = False Then
        ' The parent is not a file system folder (is more than likely a namespace
        ' extension's virtual folder), add the Type column only
        Call .ColumnHeaders.Add(, , GetResourceString("shell32.dll", IDS_SHELL32_TYPE), 120, lvwColumnLeft)
      Else
        ' Parent is file system, we'll use FindFirst/NextFile for all details
        Call .ColumnHeaders.Add(, , GetResourceString("shell32.dll", IDS_SHELL32_SIZE), 60, lvwColumnRight)
        Call .ColumnHeaders.Add(, , GetResourceString("shell32.dll", IDS_SHELL32_TYPE), 120, lvwColumnLeft)
        Call .ColumnHeaders.Add(, , GetResourceString("shell32.dll", IDS_SHELL32_MODIFIED), 120, lvwColumnLeft)
      End If
    
    End If       ' isfParent.CreateViewObject
  End With   ' objLV
  
  ' ======================================================
  ' Now fill up the ListView.
  
  ' Create an enumeration object for the parent folder.
  If SUCCEEDED(isfParent.EnumObjects(hwndOwner, _
                                                                SHCONTF_FOLDERS Or _
                                                                SHCONTF_INCLUDEHIDDEN Or _
                                                                SHCONTF_NONFOLDERS, _
                                                                ieidl)) Then
                                                                
    ' Enumerate the contents of the parent folder, obtaining
    ' the relative pidl for each item in the parent folder.
    Do While (ieidl.Next(1, pidlrelChild, 0) = NOERROR)
      
      ' Create a fully qualified pidl for the current child folder.
      pidlfqChild = CombinePIDLs(pidlfqParent, pidlrelChild)
      If pidlfqChild Then
        
        ' Add the child item to the ListView, letting the proc know what
        ' columns to use if the parent folder does not support IShellDetails
        Call InsertListItem(objLV, isfParent, pidlfqChild, pidlrelChild, isdParent, ulAttrs And SFGAO_FILESYSTEM)
        
        ' Free the current child folder's absolute pidl we created.
        isMalloc.Free ByVal pidlfqChild
      End If  ' pidlfqChild
      
      ' Free the relative pidl the enumeration gave us.
      isMalloc.Free ByVal pidlrelChild
    
    Loop   ' ieidl.Next
  End If   ' SUCCEEDED(EnumObjects))
  
'  Call ListView_SetCallbackMask(objLV.hWnd, 0)   ' LVIS_FOCUSED Or LVIS_SELECTED Or LVIS_OVERLAYMASK)
  
  ' ======================================================
  ' Finally, set the header icons and sort the items
  
  With objLV
    ' If the ListView has any items...
    If .ListItems.Count Then
      ' Set the module level variable with the selected TreeView folder's
      ' IShellFolder (is used to sort the items in ListViewCompareItems)
      Set m_isfParentFolder = isfParent
      
      ' Set the header icons
      Call Form1.HdrIcons.SetHeaderIcons(.SortKey, .SortOrder)
            
      ' Make sure the selected item is deselected (only necessary with the
      ' Mscomctl.ocx ListView)
      .SelectedItem.Selected = False
      
      ' Invoke the ListView's sort procedure via the API using the current
      ' sort order and column
      Call ListView_SortItems(.hWnd, AddressOf ListViewCompareProc, _
                                           .SortKey Or (CBool(.SortOrder) And SORT_DESCENDING))
    
      ' Let processing happen (again only for the Mscomctl.ocx ListView),
      ' select the first ListItem and make it visible (the ListView's SelectItem
      ' and ListItem Selected properties are almost completetly useless...)
      DoEvents
      Call ListView_SetFocusedItem(.hWnd, 0)
      Call ListView_EnsureVisible(.hWnd, 0, CFalse)
    End If
  End With
  
  Screen.MousePointer = vbNormal

End Sub

' Frees the pidls we held onto in InsertListItem below and empties the ListView.

Public Sub ClearListView(objLV As ListView)
  Dim li As ListItem
  Dim pidlRel  As Variant
  
  ' Free each relative pidl we stored in InsertListItem below
  For Each pidlRel In m_colPidlRels
    Call FreePIDL((pidlRel))
  Next
  Set m_colPidlRels = Nothing
  
  ' Make sure everything we allocated and freed goes back to the OS.
  Call isMalloc.HeapMinimize
  
  ' Clear the ColumnHeaders and the ListView.
  objLV.ColumnHeaders.Clear
  objLV.ListItems.Clear
  
  ' A bug: http://support.microsoft.com/support/kb/articles/q143/4/06.asp
  objLV.Arrange = lvwAutoTop
  DoEvents
  
End Sub

' Inserts the specified shell item to the specified ListView
'
'   objLV               - TreeView object reference
'   isfParent           - parent folder's IShellFolder reference
'   pidlfqChild        - pidl of the child folder being inserted, relative to the desktop folder
'   pidlrelChild        - pidl of the child folder being inserted, relative to isfParent
'   fParentIsFileSystem  - flag specifying whether the parent folder resides in the file system

' If returns the new item's real listview index.

' Called only from FillListView above

Public Function InsertListItem(objLV As ListView, _
                                              isfParent As IShellFolder, _
                                              pidlfqChild As Long, _
                                              pidlrelChild As Long, _
                                              isdParent As IShellDetails, _
                                              fParentIsFileSystem As Boolean) As Long
  Static ulAttr As Long   ' local variables are static just for performance
  Dim li As ListItem
  Dim shci As SHColInfo   ' (is actually defined in the docs as the SHELLDETAILS struct)
  Dim i As Integer
  Static hFile As Long
  Static wfd As WIN32_FIND_DATA
  Static dblKBs As Double
  Static lvi As LVITEM

  ' First get the item's shell attributes, but only those we're interested in...
  ulAttr = SFGAO_SHARE Or SFGAO_LINK Or SFGAO_FOLDER
  Call isfParent.GetAttributesOf(1, pidlrelChild, ulAttr)
  
  ' ======================================================
  ' ListItem and SubItem text

  ' Insert a ListItem for the Name column.
  Set li = objLV.ListItems.Add(Text:=GetFolderDisplayName(isfParent, pidlrelChild, SHGDN_INFOLDER))

'' debug code for understanding exactly how to interpret the
'' SHColInfo.Width value in FillListView above...
'Static pidlfqParentPrev
'Static nItems As Long
'Static cx As Long
'If (pidlfqParentPrev <> pidlfqParent) Then
'  pidlfqParentPrev = pidlfqParent
'  nItems = 0
'  cx = 0
'End If
''nItems = nItems + 1
''cx = cx + Form1.TextWidth(li)
''Debug.Print "av item.Text width: " & cx \ nItems
'If cx < Form1.TextWidth(li) Then cx = Form1.TextWidth(li)
'Debug.Print "max item.Text width: " & cx

  ' If the parent folder implements IShellDetails...
  If (isdParent Is Nothing) = False Then
    
    ' Add item's column header details per its IShellDetails info.
    For i = 1 To objLV.ColumnHeaders.Count - 1
      Call isdParent.GetDetailsOf(pidlrelChild, i, shci)
      li.SubItems(i) = GetStrRet(shci.Text, pidlrelChild)
    Next

  Else
    ' The parent folder does not implement IShellDetails, explicitly set the detatils...
    
    If (fParentIsFileSystem = False) Then
      ' The parent is not a file system folder (is more than likely a namespace
      ' extension's virtual folder), add Type column details only.
      li.SubItems(1) = GetFileTypeNamePIDL(pidlfqChild)
  
    Else
      ' Parent is file system, we'll use FindFirst/NextFile for details
      ' Type column
      li.SubItems(2) = GetFileTypeNamePIDL(pidlfqChild)
  
      ' If the current item is a file (not a folder), add Size and Modified column details
      If (ulAttr And SFGAO_FOLDER) = False Then
        ' Try to open it and get it's WIN32_FIND_DATA
        hFile = FindFirstFile(GetPathFromPIDL(pidlfqChild), wfd)
        If (hFile <> INVALID_HANDLE_VALUE) Then
  
          ' Got it, close the handle.
          Call FindClose(hFile)
  
          ' Size column (in KBs)
          dblKBs = ((wfd.nFileSizeHigh * MAXDWORD) + wfd.nFileSizeLow) / 1024
          ' Round up to the next KB.
          dblKBs = Int(dblKBs) + Abs(CBool(dblKBs - Int(dblKBs)))
          li.SubItems(1) = Format$(dblKBs, "#,##0KB") ' Size
  
          ' Modified column
          li.SubItems(3) = GetFileDateTimeStr(wfd.ftLastWriteTime)
  
        End If   '  (hFile <> INVALID_HANDLE_VALUE)
      End If   ' (ulAttr And SFGAO_FOLDER) = False
    End If   ' (fParentIsFileSystem = False)
    
  End If   ' (isdParent Is Nothing) = False

  ' ======================================================
  ' ListItem icon

  ' Set the item index (is zero-based)
  lvi.iItem = li.Index - 1

  ' Get the item's icon index within the system's small imagelist.
  ' (indices of images in the small and large imagelists are the same)
  lvi.iImage = GetFileIconIndexPIDL(pidlfqChild, SHGFI_SMALLICON)

  ' Add any overlay image...
  ' Overlay images reside in bits 8-11 of the system's normal imagelist
  ' (as opposed to a state imagelist). The share overlay is the 1st image,
  ' shortcut is 2nd, 3rd, and 4th images vary. The SFGAO_SHARE
  ' (folders) and SFGAO_LINK (files) attributes are mutually exclusive.
  If (ulAttr And (SFGAO_SHARE Or SFGAO_LINK)) Then
    lvi.mask = LVIF_IMAGE Or LVIF_STATE
    lvi.stateMask = LVIS_OVERLAYMASK
    If (ulAttr And SFGAO_SHARE) Then
      lvi.state = INDEXTOOVERLAYMASK(1)
    Else   ' (ulAttr And SFGAO_LINK)
      lvi.state = INDEXTOOVERLAYMASK(2)
    End If
  Else
    lvi.mask = LVIF_IMAGE   ' no overlay...
  End If

  ' And set the item's icon, with any overlay
  Call ListView_SetItem(objLV.hWnd, lvi)

  ' ====================================================
  
  ' Store a copy of the item's relative pidl in the collection (the ListItem's
  ' LVITEM.lParam is the collection key), the pidl is now ours, and will be
  ' freed in ClearListView above.
  Call m_colPidlRels.Add(CopyPIDL(pidlrelChild), CStr(GetLVItemlParam(objLV.hWnd, lvi.iItem)))
  
  ' Return the item's index.
  InsertListItem = lvi.iItem
  
End Function

' Returns the value of the specifed listview item's lParam.

Public Function GetLVItemlParam(hwndLV As Long, iItem As Long) As Long
  Dim lvi As LVITEM
  
  lvi.mask = LVIF_PARAM
  lvi.iItem = iItem
  If ListView_GetItem(hwndLV, lvi) Then
    GetLVItemlParam = lvi.lParam
  End If

End Function

' Returns a reference to the ListItem object residing at the specified one-based
' position within the ListView (what would assumed to be the ListItem.Index
' value if the ListView wasn't sorted using the API).

' This is a generic routine, that will resolve both Comctl32.ocx and Mscomctl.ocx ListViews.

Public Function GetListItemFromPos(objLV As ListView, iListItem As Long) As ListItem
  Dim rc As RECT

  ' Get the pixel coordinates of the specified item in the ListView. Specify that we
  ' want the item's selection rect (the item may either have no icon, or no text)
  If ListView_GetItemRect(objLV.hWnd, iListItem - 1, rc, LVIR_SELECTBOUNDS) Then

    ' Return the ListItem reference (the ListView's HitTest method *always* expects Twips)
    Set GetListItemFromPos = objLV.HitTest(rc.Left * Screen.TwipsPerPixelX, _
                                                                    rc.Top * Screen.TwipsPerPixelY)
  End If

End Function

' Processes ListView OCM_NOTIFY notification messages

Public Function DoLVNotify(hwndLV As Long, ByVal lParam As Long) As Long
  Dim nmh As NMHDR
  Dim pt As POINTAPI
  Dim lvhti As LVHITTESTINFO
  Dim i As Long

  ' Fill the NMHDR struct
  MoveMemory nmh, ByVal lParam, Len(nmh)

  Select Case nmh.code

'    Case LVN_GETDISPINFO
'      Dim lvdi As NMLVDISPINFO
'      Dim lvi As LVITEM
'
'      MoveMemory lvdi, ByVal lParam, Len(lvdi)
'If (lvdi.item.mask = LVIF_IMAGE) Then DoLVNotify = 1
'
'Debug.Print lvdi.item.mask, lvdi.item.iImage
'If (lvdi.item.mask = LVIF_IMAGE) Then
'lvi.mask = LVIF_IMAGE   ' no overlay...
'Call ListView_GetItem(hwndLV, lvi)
'lvdi.item.iImage = lvi.iImage
'MoveMemory ByVal lParam, lvdi, Len(lvdi)
'End If

    ' ======================================================
    ' Expand the selected TreeView folder, find and select the double clicked
    ' ListView folder in the TreeView, loading the ListView with that folder's contents.
    
    Case NM_DBLCLK, NM_RETURN   ' lParam = lp NMHDR
      Dim ulAtrrs As Long
      Dim pidlRel As Long
      Dim sLVItem As String
      Dim nodChild As Node
      Dim tvid As New cTVItemData
      Dim sei As SHELLEXECUTEINFO
      
      If (nmh.code = NM_DBLCLK) Then
        ' Get the index of the double clicked ListView item.
        Call GetCursorPos(pt)
        Call ScreenToClient(hwndLV, pt)
        lvhti.pt = pt
        Call ListView_HitTest(hwndLV, lvhti)
        i = lvhti.iItem
      Else   ' NM_RETURN
        ' Get the index of the selected ListView item
        i = ListView_GetSelectedItem(hwndLV)
      End If
      
      ' If we have a valid ListView item index...
      If (i <> LVI_NOITEM) Then
        
        ' Get the selected ListView item's relative pidl (we're just reading it, and
        ' must not free it)
        pidlRel = m_colPidlRels(CStr(GetLVItemlParam(hwndLV, i)))
        
        ' If the selected ListView item is a folder...
        ulAtrrs = SFGAO_FOLDER
        Call m_isfParentFolder.GetAttributesOf(1, pidlRel, ulAtrrs)
        If (ulAtrrs And SFGAO_FOLDER) Then
            
          ' Expand the selected TreeView folder, get a reference to the selected
          ' folder's first child folder, and get the Text of the selected ListView folder.
          Form1.TreeView1.SelectedItem.Expanded = True
          Set nodChild = Form1.TreeView1.SelectedItem.Child
          sLVItem = Form1.ListView1.SelectedItem
          
          ' Search under the expanded TreeView folder for the displayname of
          ' the selected ListView folder
          Do While ((nodChild Is Nothing) = False)
            If (nodChild = sLVItem) Then
              ' We found it, select it invoking a DoTVNotify/TVN_SELCHANGED
              ' filling up the ListView with its contents, we're done...
              nodChild.Selected = True
              Exit Do
            End If
            Set nodChild = nodChild.Next
          Loop
        
        Else
          ' The selected listview item is a file, execute it's associated program
          ' or command
          
          With sei
            .cbSize = Len(sei)
            .fMask = SEE_MASK_INVOKEIDLIST
            .hWnd = GetTopLevelParent(hwndLV)
'            .lpVerb = StrPtr(String$(5, 0))
'            Call lstrcpyA(ByVal .lpVerb, ByVal "Open")   ' Open is default verb
            .nShow = SW_SHOWNORMAL
            .hInstApp = App.hInstance
          End With
          
          ' Get the selected TreeView folder's item data and create a fully qualified
          ' pidl for the selected listview item
          Set tvid = GetTVItemData(Form1.TreeView1.hWnd, TreeView_GetSelection(Form1.TreeView1.hWnd))
          If (tvid Is Nothing) = False Then
            sei.lpIDList = CombinePIDLs(tvid.pidlFQ, pidlRel)
          
            If (ShellExecuteEx(sei) = 0) Then   ' rtns non-zero on success
              MsgBox "ShellExecuteEx failure" & vbCrLf & vbCrLf & _
                            "APIErr: " & Err.LastDllError & vbCrLf & GetAPIErrStr(Err.LastDllError)
            End If
              
            ' free the pidl we just created
            If sei.lpIDList Then isMalloc.Free ByVal sei.lpIDList
          End If   ' (tvid Is Nothing) = False
          
        End If   ' (ulAtrrs And SFGAO_FOLDER)
      End If   '  (i <> LVI_NOITEM)
      
    ' ======================================================
    ' Show the view context menu, or the right-clicked ListView item's shell
    ' context menu

    Case NM_RCLICK   ' lParam = lp NMHDR
      Dim nItems  As Long
      Dim apidlRels() As Long
      
      ' Get the index of the right clicked ListView item.
      Call GetCursorPos(pt)
      Call ScreenToClient(hwndLV, pt)
      lvhti.pt = pt
      Call ListView_HitTest(hwndLV, lvhti)
      
      If (lvhti.flags And LVHT_ONITEM) = False Then
        ' Not on an item, show the view context menu, and cancel the
        ' notification, hopefully ensuring that the app is out of menu mode
        ' if the menu is canceled (sometime this doesn't even work...?).
        Call Form1.PopupMenu(Form1.mnuView, vbPopupMenuLeftButton Or vbPopupMenuRightButton)
        DoLVNotify = 1
        
      Else
        ' On an item, show the item's shell context menu
        
        ' Add the pidl of the right clicked (focused) item as the first element
        ' in the pidl array, this pidl will take context menu verb precedence.
        ReDim Preserve apidlRels(0)
        apidlRels(0) = m_colPidlRels(CStr(GetLVItemlParam(hwndLV, lvhti.iItem)))
        nItems = 1
        
        ' Add the pidls of the rest of any selected items to the pidl array
        i = LVI_NOITEM
        Do
          i = ListView_GetNextItem(hwndLV, i, LVNI_SELECTED)
          If (i <> LVI_NOITEM) And (i <> lvhti.iItem) Then
            ReDim Preserve apidlRels(nItems)
            apidlRels(nItems) = m_colPidlRels(CStr(GetLVItemlParam(hwndLV, i)))
            nItems = nItems + 1
          End If
        Loop Until (i = LVI_NOITEM)
                
        ' Convert the ListView client coods back to screen coords and show
        ' the context menu for the selected item(s). If a menu command is
        ' not executed, cancel the notification...
        Call ClientToScreen(hwndLV, pt)
        If (ShowShellContextMenu(hwndLV, m_isfParentFolder, _
                                                  nItems, apidlRels(0), pt) = False) Then
          DoLVNotify = 1
        End If
      
      End If   ' (lvhti.flags And LVHT_ONITEM) = False

'    ' ======================================================
'    ' Free each deleted item's relative pidl (handled in ClearListView)
'
'    Case LVN_DELETEITEM   ' lParam = lp NMLISTVIEW
'      Dim nmlv As NMLISTVIEW
'
'      ' Fill the NMLISTVIEW struct
'      MoveMemory nmlv, ByVal lParam, Len(nmlv)
'
'      ' Get the item data from the lParam, free the item's pidls,
'      ' and remove the item data from the collection
'      isMalloc.Free ByVal CStr(nmlv.lParam)
'      Call m_colPidlRels.Remove(CStr(nmlv.lParam))
  
  End Select

End Function

' Application-defined callback function, which is called by the listview during
' a sort operation each time the relative order of two listview items needs to
' be compared. (see the desciption of LVM_SORTITEMS in the SDK)

'    lParam1      - the 1st item's LVITEM lParam value
'    lParam2      - the 2nd item's LVITEM lParam value
'    lParamSort  - application-defined value that is passed to the comparison function.

' The callback function must return a negative value if the first item should
' precede the second, a positive value if the first item should follow the second,
' or zero if the two items are equivalent.

' Invoked by ListView_SortItems() below.

Public Function ListViewCompareProc(ByVal lParam1 As Long, _
                                                            ByVal lParam2 As Long, _
                                                            ByVal lParamSort As Long) As Long
  Dim hr As Long
  
'Debug.Print GetFolderDisplayName(m_isfParentFolder, m_colPidlRels(CStr(lParam1)), SHGDN_INFOLDER), _
                    GetFolderDisplayName(m_isfParentFolder, m_colPidlRels(CStr(lParam2)), SHGDN_INFOLDER)
  
  ' Use the current parent folder's (the selected TreeView folder's)
  ' IShellFolder to compare each item, specifying the zero-based
  ' index of the column being sorted for the first param.
  hr = m_isfParentFolder.CompareIDs(lParamSort And Not SORT_DESCENDING, _
                                                          m_colPidlRels(CStr(lParam1)), _
                                                          m_colPidlRels(CStr(lParam2)))
  If (hr >= NOERROR) Then
    If (lParamSort And SORT_DESCENDING) = False Then  ' lvwAscending
      ListViewCompareProc = LOWORD(hr)
    Else
      ' lvwDescending, reverse the sign of the return value.
      ListViewCompareProc = LOWORD(hr) * -1
    End If
  End If

End Function

' ============================================================
' listview macros

Public Function ListView_SetImageList(hWnd As Long, himl As Long, iImageList As Long) As Long
  ListView_SetImageList = SendMessage(hWnd, LVM_SETIMAGELIST, ByVal iImageList, ByVal himl)
End Function
 
Public Function ListView_GetItem(hWnd As Long, pItem As LVITEM) As Boolean
  ListView_GetItem = SendMessage(hWnd, LVM_GETITEM, 0, pItem)
End Function
 
Public Function ListView_SetItem(hWnd As Long, pItem As LVITEM) As Boolean
  ListView_SetItem = SendMessage(hWnd, LVM_SETITEM, 0, pItem)
End Function
'
'Public Function ListView_SetCallbackMask(hWnd As Long, mask As Long) As Boolean
'  ListView_SetCallbackMask = SendMessage(hWnd, LVM_SETCALLBACKMASK, ByVal mask, 0)
'End Function

' ListView_GetNextItem

Public Function ListView_GetNextItem(hWnd As Long, i As Long, flags As Long) As Long
  ListView_GetNextItem = SendMessage(hWnd, LVM_GETNEXTITEM, ByVal i, ByVal flags)    ' ByVal MAKELPARAM(flags, 0))
End Function

' Returns the index of the item that is selected and has the focus rectangle (user-defined)

Public Function ListView_GetSelectedItem(hwndLV As Long) As Long
  ListView_GetSelectedItem = ListView_GetNextItem(hwndLV, -1, LVNI_FOCUSED Or LVNI_SELECTED)
End Function

Public Function ListView_GetItemRect(hWnd As Long, i As Long, prc As RECT, code As Long) As Boolean
  prc.Left = code
  ListView_GetItemRect = SendMessage(hWnd, LVM_GETITEMRECT, ByVal i, prc)
End Function

Public Function ListView_HitTest(hwndLV As Long, pinfo As LVHITTESTINFO) As Long
  ListView_HitTest = SendMessage(hwndLV, LVM_HITTEST, 0, pinfo)
End Function
 
Public Function ListView_EnsureVisible(hwndLV As Long, i As Long, fPartialOK As CBoolean) As Boolean
  ListView_EnsureVisible = SendMessage(hwndLV, LVM_ENSUREVISIBLE, ByVal i, ByVal fPartialOK)   ' ByVal MAKELPARAM(Abs(fPartialOK), 0))
End Function

Public Function ListView_SetColumnWidth(hWnd As Long, iCol As Long, cx As Long) As Boolean
  ListView_SetColumnWidth = SendMessage(hWnd, LVM_SETCOLUMNWIDTH, ByVal iCol, ByVal cx)   ' ByVal MAKELPARAM(cx, 0))
End Function

' ListView_SetItemState

Public Function ListView_SetItemState(hwndLV As Long, i As Long, state As Long, mask As Long) As Boolean
  Dim lvi As LVITEM
  lvi.state = state
  lvi.stateMask = mask
  ListView_SetItemState = SendMessage(hwndLV, LVM_SETITEMSTATE, ByVal i, lvi)
End Function

' Selects all listview items. The item with the focus rectangle maintains it (user-defined).

Public Function ListView_SelectAll(hwndLV As Long) As Boolean
  ListView_SelectAll = ListView_SetItemState(hwndLV, -1, LVIS_SELECTED, LVIS_SELECTED)
End Function
 
' Selects the specified item and gives it the focus rectangle.
' does not de-select any currently selected items (user-defined).

Public Function ListView_SetFocusedItem(hwndLV As Long, i As Long) As Boolean
  ListView_SetFocusedItem = ListView_SetItemState(hwndLV, i, LVIS_FOCUSED Or LVIS_SELECTED, _
                                                                                                    LVIS_FOCUSED Or LVIS_SELECTED)
End Function

Public Function ListView_SortItems(hwndLV As Long, pfnCompare As Long, lParamSort As Long) As Boolean
  ListView_SortItems = SendMessage(hwndLV, LVM_SORTITEMS, ByVal lParamSort, ByVal pfnCompare)
End Function

' ============================================================
' imagelist macros

' Returns the one-based index of the specified overlay image, shifted left eight bits.

Public Function INDEXTOOVERLAYMASK(iOverlay As Long) As Long
  '   INDEXTOOVERLAYMASK(i)   ((i) << 8)
  INDEXTOOVERLAYMASK = iOverlay * (2 ^ 8)
End Function
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        