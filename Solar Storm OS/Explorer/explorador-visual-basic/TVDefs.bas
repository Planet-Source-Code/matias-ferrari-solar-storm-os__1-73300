Attribute VB_Name = "mTreeviewDefs"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
'
' Procedure responsibility of pidl memory, unless specified otherwise:
' - Calling procedures are solely responsible for freeing pidls they create,
'   or receive as a return value from a called procedure.
' - Called procedures always copy pidls received in their params, and
'   *never* free pidl params.

' Contains a reference to a cTVItemData class for each Node added
' to the TreeView that holds the relative and fully qualified pidls of the
' folder represented by the Node. The string value of each Node's
' TVITEM lParam member is its respective collection key.
Private m_colTVItemData As New Collection

' ===================================================================
' treeview definitions

' messages
Public Const TV_FIRST = &H1100
Public Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Public Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Public Const TVM_GETITEM = (TV_FIRST + 12)
Public Const TVM_SETITEM = (TV_FIRST + 13)
Public Const TVM_HITTEST = (TV_FIRST + 17)
Public Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)

' TVM_GET/SETIMAGELIST wParam
Public Const TVSIL_NORMAL = 0

' TVM_GETNEXTITEM wParam
Public Const TVGN_ROOT = &H0
Public Const TVGN_NEXT = &H1
Public Const TVGN_CHILD = &H4
Public Const TVGN_CARET = &H9

' TVM_GET/SETITEM lParam
Public Type TVITEM   ' was TV_ITEM
  mask As Long
  hitem As Long
  state As Long
  stateMask As Long
  pszText As Long    ' if a string, must be pre-allocated!!
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type

' TVITEM mask
Public Const TVIF_TEXT = &H1
Public Const TVIF_IMAGE = &H2
Public Const TVIF_PARAM = &H4
Public Const TVIF_STATE = &H8
Public Const TVIF_SELECTEDIMAGE = &H20
Public Const TVIF_CHILDREN = &H40

' TVITEM state, stateMask
Public Const TVIS_EXPANDED = &H20
Public Const TVIS_EXPANDEDONCE = &H40
Public Const TVIS_OVERLAYMASK = &HF00

' TVM_HITTEST lParam
Public Type TVHITTESTINFO   ' was TV_HITTESTINFO
  pt As POINTAPI
  flags As TVHT_flags
  hitem As Long
End Type

Public Enum TVHT_flags
  TVHT_NOWHERE = &H1   ' In the client area, but below the last item
  TVHT_ONITEMICON = &H2
  TVHT_ONITEMLABEL = &H4
  TVHT_ONITEMINDENT = &H8
  TVHT_ONITEMBUTTON = &H10
  TVHT_ONITEMRIGHT = &H20
  TVHT_ONITEMSTATEICON = &H40
  TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
  
  ' user-defined
  TVHT_ONITEMLINE = (TVHT_ONITEM Or TVHT_ONITEMINDENT Or TVHT_ONITEMBUTTON Or TVHT_ONITEMRIGHT)
  
'  TVHT_ABOVE = &H100
'  TVHT_BELOW = &H200
'  TVHT_TORIGHT = &H400
'  TVHT_TOLEFT = &H800
End Enum

' TVM_SORTCHILDRENCB lParam
Public Type TVSORTCB   ' was TV_SORTCB
  hParent As Long
  lpfnCompare As Long
  lParam As Long
End Type

' notifications
Public Const TVN_FIRST = -400&   ' (0U-400U)
Public Const TVN_SELCHANGED = (TVN_FIRST - 2)         ' lParam = NMTREEVIEW
Public Const TVN_ITEMEXPANDING = (TVN_FIRST - 5)    ' lParam = NMTREEVIEW
Public Const TVN_DELETEITEM = (TVN_FIRST - 9)           ' lParam = NMTREEVIEW

' lParam for most treeview notification messages
Public Type NMTREEVIEW   ' was NM_TREEVIEW
  hdr As NMHDR
  ' Specifies a notification-specific action flag.
  ' Is TVC_* for TVN_SELCHANGING, TVN_SELCHANGED, TVN_SETDISPINFO
  ' Is TVE_* for TVN_ITEMEXPANDING, TVN_ITEMEXPANDED
  action As Long
  itemOld As TVITEM
  itemNew As TVITEM
  ptDrag As POINTAPI
End Type
'

' Inserts the specified root folder into the TreeView control.
'
'   objTV   - TreeView object reference
'   pidlFQ  - root folder's fully qualified pidl, is not copied, must not be freed by caller.
'
' If successful, returns the root folder's treeview item handle, returns 0 otherwise.
'
' Called from Form1.Form_Load and Form1.mnuFileRoot_Click

Public Function InsertRootFolder(objTV As TreeView, pidlFQ As Long) As Long
  Dim pidlRel As Long
  Dim hitem As Long
  
  Call RemoveRootFolder(objTV)
  
  ' Get the specified root folder's relative pidl (we have to free it,
  ' InsertFolder copies it).
  pidlRel = GetItemID(pidlFQ, GIID_LAST)
  If pidlRel Then
    
    ' Insert the root folder (both pidls are copied in the call)
    hitem = InsertFolder(objTV, Nothing, GetIShellFolderParent(pidlFQ), pidlFQ, pidlRel, 0, 0)
    If hitem Then
    
      ' Select and expand the root, the latter invoking a TVN_ITEMEXPANDING,
      ' which calls InsertSubfolders() loading subfolders under the root
      objTV.Nodes(1).Selected = True
      objTV.Nodes(1).Expanded = True
      
      InsertRootFolder = hitem
    End If   ' hitem
  
    Call FreePIDL(pidlRel)
  End If   ' pidlRel

End Function

' Removes the root folder and all of its subfolders from the specified TreeView,

' called from InsertRootFolder above and Form1.Form_Unload

Public Sub RemoveRootFolder(objTV As TreeView)
  
  If objTV.Nodes.Count Then
    ' Collapse the root folder and remove it.
    objTV.Nodes(1).Root.Expanded = False
    
    ' Invokes a DoTVNotify/TVN_DELETEITEM freeing the pidls
    ' we stored in InsertFolder below (we do not want to iterate the
    ' m_colTVItemData collection here and free all pidls since
    ' RefreshTreeview may remove some, but not all, treeview items).
    Call objTV.Nodes.Remove(objTV.Nodes(1).Root.Index)
  End If
  
End Sub

' Inserts the specified folder under the specified parent folder
'
'   objTV               - TreeView object reference
'   nodParent        - parent folder's Node reference, is Nothing for the root Node
'   isfParent           - parent folder's IShellFolder reference
'   pidlfqChild        - pidl of the child folder being inserted, relative to the desktop folder
'   pidlrelChild        - pidl of the child folder being inserted, relative to isfParent
'   hitemParent      - parent folder's treeview item handle, is 0 for root folder
'   hitemPrevChild - parent folder's previous child's treeview item handle, is 0 for parent's first child

' If successful, returns the folder's treeview item handle, returns 0 otherwise.

' Called from InsertRootFolder above, and InsertSubfolders below.

Public Function InsertFolder(objTV As TreeView, _
                                            nodParent As Node, _
                                            isfParent As IShellFolder, _
                                            pidlfqChild As Long, _
                                            pidlrelChild As Long, _
                                            hitemParent As Long, _
                                            hitemPrevChild As Long) As Long
  Dim ulAttrs As ESFGAO
  Dim tvi As TVITEM
  Dim tvid As New cTVItemData
  
  ' Get the child folder's attributes, specifiy what attributes we want.
  ulAttrs = SFGAO_HASSUBFOLDER Or SFGAO_SHARE
  Call isfParent.GetAttributesOf(1, pidlrelChild, ulAttrs)
  
  ' ====================================================
  ' Fill the folder's TVITEM struct
  
  ' By explicitly setting the treeview item attributes that the VB TreeView
  ' normally does callbacks for, we increase the performance of the
  ' TreeView dramatically, since a TVN_GETDISPINFO is sent by the real
  ' treeview any time, any item, needs to present any of these attributes.
  ' One problem though, this information is not available in the Node...
  ' ...an insignificant side effect since this info can be obtained by APIs...
  
  ' Indicate what TVITEM members will contain data
  tvi.mask = TVIF_CHILDREN Or TVIF_IMAGE Or TVIF_SELECTEDIMAGE
  
  ' If the folder has subfolders, explicitly give the item a button, overriding
  ' the I_CHILDRENCALLBACK value the VB TreeView uses.
  tvi.cChildren = Abs(CBool(ulAttrs And SFGAO_HASSUBFOLDER))
  
  ' Explicitly set the folder's normal and selected icon indices, overriding
  ' the I_IMAGECALLBACK value the VB TreeView uses.
  tvi.iImage = GetFileIconIndexPIDL(pidlfqChild, SHGFI_SMALLICON)
  tvi.iSelectedImage = GetFileIconIndexPIDL(pidlfqChild, SHGFI_SMALLICON Or SHGFI_OPENICON)

  ' If the folder is shared, give it the share overlay icon.
  If (ulAttrs And SFGAO_SHARE) Then
    tvi.mask = tvi.mask Or TVIF_STATE
    tvi.state = TVIS_OVERLAYMASK
    ' share overlay is the 1st system imagelist overlay image,
    ' shortcut is 2nd, gray arrow is 3rd, no 4th image
    tvi.stateMask = INDEXTOOVERLAYMASK(1)
  End If
  
  ' ====================================================
  
  ' Add the Node to the TreeView, without a button or icons (we did
  ' everything above first so that there's the least amount of code between
  ' inserting the Node and setting calling TreeView_SetItem below).
  If (nodParent Is Nothing) Then
    Call objTV.Nodes.Add(Text:=GetFolderDisplayName(isfParent, pidlrelChild, SHGDN_INFOLDER))
  Else
    Call objTV.Nodes.Add(nodParent, tvwChild, Text:=GetFolderDisplayName(isfParent, pidlrelChild, SHGDN_INFOLDER))
  End If
    
  ' Get the new Node's hItem
  If (hitemParent = 0) Then
    tvi.hitem = TreeView_GetRoot(objTV.hWnd)
  ElseIf (hitemPrevChild = 0) Then
    tvi.hitem = TreeView_GetChild(objTV.hWnd, hitemParent)
  Else
    tvi.hitem = TreeView_GetNextSibling(objTV.hWnd, hitemPrevChild)
  End If
  
  ' And set the item's button and icons, done deal...
  Call TreeView_SetItem(objTV.hWnd, tvi)
  
  ' ====================================================
  
  ' Finally, copy the pidls and add them to the item data collection, they
  ' will eventually be freed in DoTVNotify/TVN_DELETEITEM below.
  tvid.pidlFQ = CopyPIDL(pidlfqChild)
  tvid.pidlRel = CopyPIDL(pidlrelChild)
  Call m_colTVItemData.Add(tvid, CStr(GetTVItemlParam(objTV.hWnd, tvi.hitem)))
  
  ' Return the folder's hItem
  InsertFolder = tvi.hitem

End Function

' Inserts subfolders under the specified TreeView's parent folder.

'   objTV            - TreeView object reference
'   pidlfqParent   - parent folder's fully qualified pidl, relative to the desktop
'   hitemParent   - parent folder's treeview item handle
'   nodParent     - parent folder's Node reference

' Called only from DoTVNotify/TVN_ITEMEXPANDING

Public Sub InsertSubfolders(objTV As TreeView, _
                                            pidlfqParent As Long, _
                                            hitemParent As Long, _
                                            nodParent As Node)
  Dim hwndOwner As Long
  Dim isfParent As IShellFolder         ' parent folder's IShellFolder
  Dim ieidl As IEnumIDList                ' isfParent's enumeration object
  Dim pidlrelChild As Long                ' child folder's pidl, relative to isfParent
  Dim pidlfqChild As Long                ' child folder's fully qualified pidl, relative to the desktop folder
  Dim hitemChild As Long
  Dim tvscb As TVSORTCB
  Dim tvi As TVITEM
  
  Screen.MousePointer = vbHourglass
  hwndOwner = GetTopLevelParent(objTV.hWnd)
  
  ' Get the parent's IShellFolder from its fully qualified pidl
  Set isfParent = GetIShellFolder(isfDesktop, pidlfqParent)
  
  ' Create an enumeration object for the parent folder.
  If SUCCEEDED(isfParent.EnumObjects(hwndOwner, _
                                                                SHCONTF_FOLDERS Or SHCONTF_INCLUDEHIDDEN, _
                                                                ieidl)) Then
                                                                
    ' Enumerate the contents of the parent folder
    Do While (ieidl.Next(1, pidlrelChild, 0) = NOERROR)
      
      ' Create a fully qualified pidl for the current child folder.
      pidlfqChild = CombinePIDLs(pidlfqParent, pidlrelChild)
      If pidlfqChild Then
      
        ' Insert the child folder under the parent folder.
        hitemChild = InsertFolder(objTV, nodParent, isfParent, pidlfqChild, pidlrelChild, hitemParent, hitemChild)
        
        ' Free the current child folder's absolute pidl we created.
        isMalloc.Free ByVal pidlfqChild
      End If  ' pidlfqChild
      
      ' Free the relative pidl the enumeration gave us.
      isMalloc.Free ByVal pidlrelChild
    
    Loop   ' ieidl.Next
  End If   ' SUCCEEDED(EnumObjects))
    
  If hitemChild Then
    ' Setup the callback and sort the parent folder
    tvscb.hParent = hitemParent
'    tvscb.lpfnCompare = FARPROC(AddressOf TreeViewCompareProc)
    MoveMemory tvscb.lpfnCompare, AddressOf TreeViewCompareProc, 4
    tvscb.lParam = ObjPtr(isfParent)
    Call TreeView_SortChildrenCB(objTV.hWnd, tvscb, 0)
  
  Else
    ' The parent folder is expanding yet it has no subfolders. We'll assume
    ' that we've encountered a network folder which ISF::GetAttributesOf
    ' couldn't resolve when the folder was originally inserted. Remove its button
    tvi.hitem = hitemParent
    tvi.mask = TVIF_CHILDREN
    tvi.cChildren = 0
    Call TreeView_SetItem(objTV.hWnd, tvi)
  End If
  
  Screen.MousePointer = vbDefault

End Sub

' Refreshes the TreeView only by removing all subfolders under all collapsed
' parent folders (a full blown treeview-folder refresh algorithm requires many
' KBs of code, and is well beyond the scope of this demo...)

'   objTV         - TreeView object reference
'   nodSibling  - Node reference of the first sibling under any given parent Node,
'                        on first call pass TreeView.Nodes(1).Root

' called only from Form1.mnuViewRefresh_Click

Public Sub RefreshTreeview(objTV As TreeView, nodSibling As Node)
  Dim nodChild  As Node
  
  Do While (nodSibling Is Nothing) = False

    ' remove all children of collapsed sibling Nodes
    If nodSibling.Expanded Then
      Call RefreshTreeview(objTV, nodSibling.Child)
    Else
      ' nodSibling.Children calls TVM_GETNEXTITEM/TVGN_NEXTs for the whole
      ' sibling hierarchy, this method sends the least amount of TVM_GETNEXTITEMs...
      ' And be sure that the parent Node Sorted = False before re-inserting children...
      Set nodChild = nodSibling.Child
      Do While (nodChild Is Nothing) = False
        objTV.Nodes.Remove nodChild.Index
        Set nodChild = nodSibling.Child
      Loop
    End If

    Set nodSibling = nodSibling.Next
  Loop

End Sub

' Returns the lParam of the specified treeview item.

Public Function GetTVItemlParam(hwndTV As Long, hitem As Long) As Long
  Dim tvi As TVITEM
  
  tvi.hitem = hitem
  tvi.mask = TVIF_PARAM
  
  If TreeView_GetItem(hwndTV, tvi) Then
    GetTVItemlParam = tvi.lParam
  End If

End Function

  ' Returns the specified item's item data from it's lParam.

'   hwndTV       - treeview's window handle
'   hItem            - treeview item's handle

' Called only from Form1.mnuFileRoot_Click

Public Function GetTVItemData(hwndTV As Long, hitem As Long) As cTVItemData
  Set GetTVItemData = m_colTVItemData(CStr(GetTVItemlParam(hwndTV, hitem)))
End Function
'
'Public Function GetNodeFromhItem(hwndTV As Long, hItem As Long) As Node
'  Set GetNodeFromhItem = GetNodeFromlParam(GetTVItemlParam(hwndTV, hItem))
'End Function

' Returns an AddRef'd Node object reference from the Node's
' TVITEM lParam value.

' ======================================================
' For both the Mscomctl.ocx and Comctl32.ocx TreeView and ListView
' controls, the Node and ListItem's ObjPtr() values reside at the 3rd
' DWORD (@ byte offset 8) in the Node and ListItem's lParam.
'
' Is highly undocumented, for more info, see the TVItemData demo at
' http://www.mvps.org/btmtz/treeview/
' =====================================================

Public Function GetNodeFromlParam(lParam As Long) As Node
  Dim pNode As Long
  Dim nod As Node
  
  If lParam Then
    MoveMemory pNode, ByVal lParam + 8, 4
    If pNode Then
      MoveMemory nod, pNode, 4
      Set GetNodeFromlParam = nod
      ' nod is not AddRef'd, so we have to zero it before this proc ends,
      ' or it will be Released (causing a GPF) when it goes out of scope.
      FillMemory nod, 4, 0
    End If
  End If
  
End Function

' Processes TreeView OCM_NOTIFY notification messages

Public Function DoTVNotify(hwndTV As Long, ByVal lParam As Long) As Long
  Dim nmtv As NMTREEVIEW
  Dim tvid As cTVItemData
  
  ' Fill the NMTREEVIEW struct (all members are Longs).
  ' For all WM_NOTIFY messages, lParam always points to a struct which is
  '  either the NMHDR struct itself, or whose 1st member is the NMHDR struct.
  MoveMemory nmtv, ByVal lParam, Len(nmtv)
  
  Select Case nmtv.hdr.code
    
    ' ======================================================
    ' Fill the TreeView with the children folders under the expanding parent folder
    
    Case TVN_ITEMEXPANDING   ' lParam = lp NMTREEVIEW
      
      ' If we have not already added subfolders...
'      If ((nmtv.itemNew.state And TVIS_EXPANDEDONCE) = False) Then
      If (TreeView_GetChild(hwndTV, nmtv.itemNew.hitem) = 0) Then
        Call InsertSubfolders(Form1.TreeView1, _
                                         m_colTVItemData(CStr(nmtv.itemNew.lParam)).pidlFQ, _
                                         nmtv.itemNew.hitem, _
                                         GetNodeFromlParam(nmtv.itemNew.lParam))
      End If
      
    ' ======================================================
    ' Fill the ListView with the children folders of the selected TreeView folder
    
    Case TVN_SELCHANGED   ' lParam = lp NMTREEVIEW
      
      ' Get the selected folder's item data
      Set tvid = m_colTVItemData(CStr(nmtv.itemNew.lParam))
      If (tvid Is Nothing) = False Then
        Call FillListView(Form1.ListView1, tvid.pidlFQ, tvid.pidlRel)
      End If
    
    ' ======================================================
    ' Show the right-clicked treeview folder's shell context menu

    Case NM_RCLICK   ' lParam = lp NMHDR
      Dim pt As POINTAPI
      Dim tvhti As TVHITTESTINFO

      ' Get the highlighted treeview item (instead of just TVHT_ONITEM)
      Call GetCursorPos(pt)
      Call ScreenToClient(hwndTV, pt)
      tvhti.pt = pt
      Call TreeView_HitTest(hwndTV, tvhti)
      If (tvhti.flags And TVHT_ONITEMLINE) Then

        ' Get the item's item data from it's lParam
        Set tvid = m_colTVItemData(CStr(GetTVItemlParam(hwndTV, tvhti.hitem)))
        If (tvid Is Nothing) = False Then
        
          ' This test is only needed for the desktop folder
          ' (the desktop has no context menu).
          If (IsDesktopPIDL(tvid.pidlFQ) = False) Then
            ' Convert the treeview client coods back to screen coords and show
            ' the context menu for the selected item(s). If a menu command is
            ' not executed, cancel the notification...
            Call ClientToScreen(hwndTV, pt)
            If (ShowShellContextMenu(hwndTV, GetIShellFolderParent(tvid.pidlFQ), _
                                                      1, tvid.pidlRel, pt) = False) Then
              DoTVNotify = 1
            End If
          End If
                
        End If   ' (tvid Is Nothing) = False
      End If   ' (tvhti.flags And TVHT_ONITEMLINE)
      
    ' ======================================================
    ' Free each deleted item's item data pidls

    Case TVN_DELETEITEM   ' lParam = lp NMTREEVIEW

      ' Get the item data from the lParam, free the item's pidls,
      ' and remove the item data from the collection
      Set tvid = m_colTVItemData(CStr(nmtv.itemOld.lParam))
      If (tvid Is Nothing) = False Then
        isMalloc.Free ByVal tvid.pidlRel
        isMalloc.Free ByVal tvid.pidlFQ
        Call m_colTVItemData.Remove(CStr(nmtv.itemOld.lParam))
      End If
      
  End Select

End Function

' Application-defined callback function, which is called during a sort operation each time
' the relative order of two treeview items needs to be compared. Implements a Merge sort.
' (see the TVSORTCB struct's lpfnCompare member desciption in the SDK)

' The lParam1 and lParam2 parameters correspond to the lParam member of the TVITEM
' structure for the two items being compared.

'    lParam1     - pointer to the 1st item's TVITEM lParam member
'    lParam2     - pointer to the 2nd item's TVITEM lParam member
'    lParamSort - corresponds to the lParam member of the TVSORTCB structure
'                        that was passed with the TVM_SORTCHILDRENCB message.

' The callback function must return a negative value if the first item should precede the second,
' a positive value if the first item should follow the second, or zero if the two items are equivalent.

' Invoked by a TreeView_SortChildrenCB call in the InsertSubfolders proc above.

Public Function TreeViewCompareProc(ByVal lParam1 As Long, _
                                                              ByVal lParam2 As Long, _
                                                              ByVal lParamSort As Long) As Long
  Dim isfParent As IShellFolder
  Dim hr As Long   ' HRESULT
  
  ' Get the parent folder's un-AddRef'd IShellFolder
  ' from lParamSort that we set in InsertSubfolders.
  MoveMemory isfParent, lParamSort, 4

'Debug.Print GetFolderDisplayName(isfParent, m_colTVItemData(CStr(lParam1)).pidlRel, SHGDN_INFOLDER), _
                    GetFolderDisplayName(isfParent, m_colTVItemData(CStr(lParam2)).pidlRel, SHGDN_INFOLDER)
  
  hr = isfParent.CompareIDs(0, m_colTVItemData(CStr(lParam1)).pidlRel, _
                                                m_colTVItemData(CStr(lParam2)).pidlRel)

'  If SUCCEEDED(hr) Then TreeViewCompareProc = LOWORD(hr)
  If (hr >= NOERROR) Then TreeViewCompareProc = LOWORD(hr)
  
  ' Zero the IShellfolder object variable so it is not Released.
  FillMemory isfParent, 4, 0

End Function

' ===================================================================
' treeview macros

' Sets the normal or state image list for a tree-view control and redraws the control using the new images.
' Returns the handle to the previous image list, if any, or 0 otherwise.

Public Function TreeView_SetImageList(hWnd As Long, himl As Long, iImage As Long) As Long
  TreeView_SetImageList = SendMessage(hWnd, TVM_SETIMAGELIST, ByVal iImage, ByVal himl)
End Function

' TreeView_GetNextItem

' Retrieves the tree-view item that bears the specified relationship to a specified item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextItem(hWnd As Long, hitem As Long, flag As Long) As Long
  TreeView_GetNextItem = SendMessage(hWnd, TVM_GETNEXTITEM, ByVal flag, ByVal hitem)
End Function

' Retrieves the first child item. The hitem parameter must be NULL.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetChild(hWnd As Long, hitem As Long) As Long
  TreeView_GetChild = TreeView_GetNextItem(hWnd, hitem, TVGN_CHILD)
End Function

' Retrieves the next sibling item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetNextSibling(hWnd As Long, hitem As Long) As Long
  TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hitem, TVGN_NEXT)
End Function

' Retrieves the currently selected item.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetSelection(hWnd As Long) As Long
  TreeView_GetSelection = TreeView_GetNextItem(hWnd, 0, TVGN_CARET)
End Function

' Retrieves the topmost or very first item of the tree-view control.
' Returns the handle to the item if successful or 0 otherwise.

Public Function TreeView_GetRoot(hWnd As Long) As Long
  TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
End Function

' Retrieves some or all of a tree-view item's attributes.
' Returns TRUE if successful or FALSE otherwise.

Public Function TreeView_GetItem(hWnd As Long, pItem As TVITEM) As Boolean
  TreeView_GetItem = SendMessage(hWnd, TVM_GETITEM, 0, pItem)
End Function

' Sets some or all of a tree-view item's attributes.
' Old docs say returns zero if successful or - 1 otherwise.
' New docs say returns TRUE if successful, or FALSE otherwise

Public Function TreeView_SetItem(hWnd As Long, pItem As TVITEM) As Boolean
  TreeView_SetItem = SendMessage(hWnd, TVM_SETITEM, 0, pItem)
End Function

' Determines the location of the specified point relative to the client area of a tree-view control.
' Returns the handle to the tree-view item that occupies the specified point or NULL if no item
' occupies the point.

Public Function TreeView_HitTest(hWnd As Long, lpht As TVHITTESTINFO) As Long
  TreeView_HitTest = SendMessage(hWnd, TVM_HITTEST, 0, lpht)
End Function

' Sorts tree-view items using an application-defined callback function that compares the items.
' Returns TRUE if successful or FALSE otherwise.
' fRecurse is reserved for future use and must be zero.

Public Function TreeView_SortChildrenCB(hWnd As Long, psort As TVSORTCB, fRecurse As Boolean) As Boolean
  TreeView_SortChildrenCB = SendMessage(hWnd, TVM_SORTCHILDRENCB, ByVal fRecurse, psort)
End Function
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     