VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Explorador de Solar Storm"
   ClientHeight    =   6225
   ClientLeft      =   1320
   ClientTop       =   1665
   ClientWidth     =   9465
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9465
   Begin ComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   5040
      TabIndex        =   2
      Top             =   1140
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5424
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3075
      Left            =   720
      TabIndex        =   1
      Top             =   1170
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   5424
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   3900
      TabIndex        =   0
      Top             =   1380
      Width           =   435
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuFileRoot 
         Caption         =   "&Cambiar Ruta..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "S&alir"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Vistas"
      Begin VB.Menu mnuViewLargeIcons 
         Caption         =   "Iconos &Grandes"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "Iconos &Pequeños"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&Lista"
      End
      Begin VB.Menu mnuViewReport 
         Caption         =   "&Detalles"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrange 
         Caption         =   "Orden de los Iconos"
         Begin VB.Menu mnuViewArrangeAZ 
            Caption         =   "A&scendente"
         End
         Begin VB.Menu mnuViewArrangeZA 
            Caption         =   "&Descendente"
         End
         Begin VB.Menu mnuViewArrangeSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewArrangeAuto 
            Caption         =   "&Orden Automatico"
         End
      End
      Begin VB.Menu mnuViewSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSelAll 
         Caption         =   "Seleccionar &Todos"
      End
      Begin VB.Menu mnuViewSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refrescar"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' ========================================================
' This project uses subclassing, and utilizes the services of the "Debug
' Object for AddressOf Subclassing" ActiveX server, Dbgwproc.dll, which
' allows unencumbered code execution when stepping through code in
' the VB IDE. This server is freely distributable and can be obtained from
' Microsoft at http://msdn.microsoft.com/vbasic/downloads/controls.asp.

' Set the conditional compilation argument:   DEBUGWINDOWPROC = 1
' in the project properties dialog/Make tab to enable the server's services.
' ========================================================

Private m_hwndTV As Long   ' TreeView1.hWnd
Private m_hwndLV As Long   ' ListView1.hWnd

Private m_cVSplitter As New cVSplitter
Private m_cHdrIcons As New cLVHeaderSortIcons
'

Private Sub Form_Load()
  Dim dwStyle As Long
  Dim pidlDesktop As Integer  ' desktop's pidl, the TreeView's root (see below)
  
  Screen.MousePointer = vbHourglass
                                  
  ' Initialize the imagelist with an undoc shell call. Is only necessary for
  ' NT4 where the app gets an uninitialized system imagelist copy. See:
  ' http://www.geocities.com/SiliconValley/4942/iconcache.html
  ' **** Call is not exported in stock Win95's Shell32.dll !!! ****
  On Error Resume Next
  Call FileIconInit(True)
  On Error GoTo 0

  Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
  
  ' Initialize the header sort icons object.
  Set m_cHdrIcons.ListView = ListView1
  
  ' ======================================================
  ' Splitter, required property settings:
  
  KeyPreview = True
  ScaleMode = vbPixels   ' also affects the TreeView's Indentation property

'  Frame1.Appearance = 0   ' flat
'  Frame1.BorderStyle = vbBSNone   ' hides the caption
'  Frame1.ClipControls = False
  
  With m_cVSplitter
    Call .SetControls(Me, TreeView1, ListView1, Frame1)
    .Left = (ScaleWidth - .Width) * 0.3
    .TrackSplit = True
  End With
    
  ' ======================================================
  ' ListView
    
  ' Have to initialize the ListView before the TreeView is loaded below
  ' with the InsertRootFolder call (which loads the ListView).
  
  With ListView1
'    .BorderStyle = vbBSNone
    .HideSelection = False
    .LabelEdit = lvwManual
    .MultiSelect = True
    m_hwndLV = .hWnd
  End With

  ' First tell the ListView that it will share the imagelist assigned to it, so that
  ' the ListView does not destroy the imagelist when it is itself destroyed.
  dwStyle = GetWindowLong(m_hwndLV, GWL_STYLE)
  If ((dwStyle And LVS_SHAREIMAGELISTS) = False) Then
    dwStyle = dwStyle Or LVS_SHAREIMAGELISTS
  End If
  dwStyle = dwStyle And Not WS_BORDER
'  dwStyle = dwStyle And Not WS_CLIPCHILDREN
  Call SetWindowLong(m_hwndLV, GWL_STYLE, dwStyle)
  
  ' Assign the handles of the system's small and large icon imagelists to the
  ' ListView. We will set the ListItem image indices directly in LoadIcons
  ' proc below. As far as the VB ListView's internal code is concerned, it's
  ' not using an imagelist, and its ListItem icon properties return 0.
  Call ListView_SetImageList(m_hwndLV, GetSystemImagelist(SHGFI_SMALLICON), LVSIL_SMALL)
  Call ListView_SetImageList(m_hwndLV, GetSystemImagelist(SHGFI_LARGEICON), LVSIL_NORMAL)
  
  ' We need to subclass the ListView not only to prevent it from removing
  ' our system imagelist assignments (which it will do if left unchecked...),
  ' but to also catch notification messages reflected back to it from its
  ' hidden parent window via the OCM_NOTIFY ActiveX control message.
  Call SubClass(m_hwndLV, AddressOf LVWndProc)
  
  ' ======================================================
  ' TreeView
  
  ' Initialize...
  With TreeView1
    .HideSelection = False
    .Indentation = 19   ' default common control treeview indentation.
    .LabelEdit = tvwManual
    m_hwndTV = .hWnd
  End With
  
  ' Get the handle to the system's small icon imagelist and assign it to the
  ' TreeView. We will set the Node image indices directly in InsertFolder.
  ' As far as the VB TreeView's internal code is concerned, it's not using
  ' an imagelist. (treeviews share imagelists by default)
  Call TreeView_SetImageList(m_hwndTV, GetSystemImagelist(SHGFI_SMALLICON), TVSIL_NORMAL)
  
  ' We need to subclass the TreeView not only to prevent it from removing
  ' our system imagelist assignment (which it will do if left unchecked...),
  ' but to also catch notification messages reflected back to it from its
  ' hidden parent window via the OCM_NOTIFY ActiveX control message.
  Call SubClass(m_hwndTV, AddressOf TVWndProc)
    
  ' Insert the desktop folder as the TreeView root, expanding the root
  ' folder to show its subfolders (the desktop pidl is nothing more than
  ' a pointer to a NULL WORD, see the IsDesktopPIDL function)
  Call InsertRootFolder(TreeView1, VarPtr(pidlDesktop))
  
  ' ======================================================
  
  ' Set the menus and the ListView's view
  Call SwitchView(lvwReport)
  Call mnuViewArrangeAZ_Click
  
  Screen.MousePointer = vbDefault
  
End Sub

' Clean up...

Private Sub Form_Unload(Cancel As Integer)
  
  Screen.MousePointer = vbHourglass
  
  ' Hide the Form while we clear the TreeView and ListView.
  Visible = False
  
  ' Make sure all of the pidls that were stored in InsertListItem are freed.
  Call ClearListView(ListView1)
  Call UnSubClass(m_hwndLV)
  
  ' Detach the system imagelists from the ListView after we're unsubclassed
  ' It is only necessary to remove the system imagelist associations
  ' from the ListView if its LVS_SHAREIMAGELISTS style is not set.
  ' If we didn't do this, the ListView would destroy both imagelists
  ' (and that's why processes get system imagelist copies on NT)
  Call ListView_SetImageList(m_hwndLV, 0, LVSIL_SMALL)
  Call ListView_SetImageList(m_hwndLV, 0, LVSIL_NORMAL)
  
  ' Make sure all of the pidls that were stored in InsertFolder are freed.
  ' (we still need to process DoTVNotify/TVN_DELETEITEM so that
  ' all pidls held by the m_colTVItemData collection are freed).
  Call RemoveRootFolder(TreeView1)
  Call UnSubClass(m_hwndTV)
  
  ' Finally, detach the system imagelist from the TreeView after we're
  ' unsubclassed (...again unecessary, but prudent)
  Call TreeView_SetImageList(m_hwndTV, 0, TVSIL_NORMAL)

  Screen.MousePointer = vbDefault
  
End Sub

Public Property Get HdrIcons() As cLVHeaderSortIcons
  Set HdrIcons = m_cHdrIcons
End Property

' Toggles the sort order, and sorts the ListView's items or subitems under
' the respectively clicked ColumnHeader.

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
  Dim i As Integer
  
  With ListView1
    ' Toggle the clicked column's sort order only if the active colum is clicked
    ' (iow, don't reverse the sort order when different columns are clicked).
    If (.SortKey = ColumnHeader.Index - 1) Then
      ColumnHeader.Tag = Not Val(ColumnHeader.Tag)
    End If
    
    ' Set sort order to that of the respective SortOrderConstants value
    .SortOrder = Abs(Val(ColumnHeader.Tag))
    
    ' Get the zero-based index of the clicked column.
    ' (ColumnHeader.Index is one-based).
    .SortKey = ColumnHeader.Index - 1
  End With
  
  Call SortListview
  
End Sub

Private Sub SortListview()
    
  With ListView1
    ' Set the header icons
    Call m_cHdrIcons.SetHeaderIcons(.SortKey, .SortOrder)
    
    ' Sort the ListView, passing the zero-based column header index and
    ' the sort order as a flag (&H80000000) to be used by CompareIDs
    Call ListView_SortItems(m_hwndLV, AddressOf ListViewCompareProc, _
                                         .SortKey Or (CBool(.SortOrder) And SORT_DESCENDING))
  End With

End Sub

' Ctrl+A:          select all ListView items
' Backspace: select the parent folder of the TreeView's selected folder

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
  If (Shift And vbCtrlMask) And (KeyCode = vbKeyA) Then
    Call ListView_SelectAll(m_hwndLV)
  ElseIf (KeyCode = vbKeyBack) Then
    On Error Resume Next
    TreeView1.SelectedItem.Parent.Selected = True
  End If
End Sub

' ===============================================================
' File menu

Private Sub mnuFileRoot_Click()
  Dim tvid As cTVItemData
  Dim pidlfqOldRoot As Long
  Dim pidlfqNewRoot As Long
  
  ' Get the existing root folder's item data, fully qualified pidl
  ' and release the item data (pidlfqOldRoot will be freed and
  ' invalid after RemoveRootFolder is called in InsertRootFolder).
  Set tvid = GetTVItemData(m_hwndTV, TreeView_GetRoot(m_hwndTV))
  If (tvid Is Nothing) = False Then
    pidlfqOldRoot = tvid.pidlFQ
    Set tvid = Nothing
  End If
  
  pidlfqNewRoot = BrowseDialog(hWnd, "Select the TreeView root folder", 0, 0, pidlfqOldRoot)
  
  ' If a folder was selected in the dialog, set it as the new root
  ' folder, and free its pidl (it was copied).
  If pidlfqNewRoot Then
    Screen.MousePointer = vbHourglass
    Call InsertRootFolder(TreeView1, pidlfqNewRoot)
    isMalloc.Free ByVal pidlfqNewRoot
    Screen.MousePointer = vbDefault
  End If
  
End Sub

Private Sub mnuFileAbout_Click()

End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

' ===============================================================
' View menu

Private Sub mnuView_Click()
  mnuViewArrangeAuto.Enabled = ((ListView1.View = lvwIcon) Or (ListView1.View = lvwSmallIcon))
End Sub

Private Sub mnuViewLargeIcons_Click()
  Call SwitchView(lvwIcon)
End Sub

Private Sub mnuViewSmallIcons_Click()
  Call SwitchView(lvwSmallIcon)
End Sub

Private Sub mnuViewList_Click()
  Call SwitchView(lvwList)
End Sub

Private Sub mnuViewReport_Click()
  Call SwitchView(lvwReport)
End Sub

Private Sub SwitchView(dwNewView As ListViewConstants)
  
  ListView1.View = dwNewView
  
  ' A bug: http://support.microsoft.com/support/kb/articles/q143/4/06.asp
  ListView1.Arrange = lvwAutoTop
  If (mnuViewArrangeAuto.Checked = False) Then
    ListView1.Arrange = lvwNone
  End If
  
  mnuViewLargeIcons.Checked = (dwNewView = lvwIcon)
  mnuViewSmallIcons.Checked = (dwNewView = lvwSmallIcon)
  mnuViewList.Checked = (dwNewView = lvwList)
  mnuViewReport.Checked = (dwNewView = lvwReport)

End Sub

Private Sub mnuViewSelAll_Click()
  Call ListView_SelectAll(m_hwndLV)
End Sub

Private Sub mnuViewArrangeAZ_Click()
  mnuViewArrangeAZ.Checked = True
  mnuViewArrangeZA.Checked = False
  ListView1.SortOrder = lvwAscending
  Call SortListview
End Sub

Private Sub mnuViewArrangeZA_Click()
  mnuViewArrangeAZ.Checked = False
  mnuViewArrangeZA.Checked = True
  ListView1.SortOrder = lvwDescending
  Call SortListview
End Sub

Private Sub mnuViewArrangeAuto_Click()
  If (mnuViewArrangeAuto.Checked = False) Then
    mnuViewArrangeAuto.Checked = True
    ListView1.Arrange = lvwAutoTop
  Else
    mnuViewArrangeAuto.Checked = False
    ListView1.Arrange = lvwNone
  End If
End Sub

Private Sub mnuViewRefresh_Click()   ' F5
  Dim tvid As cTVItemData
  Dim pidlFQ As Long
  Dim pidlRel As Long
  
  Screen.MousePointer = vbHourglass
  
  ' Only removes child Nodes under collapsed parent Nodes.
  Call RefreshTreeview(TreeView1, TreeView1.Nodes(1))
  
  ' Get the selected folder's item data, copy the pidls (the originals will be
  ' freed by the ClearListView call below), and release the item data.
  Set tvid = GetTVItemData(m_hwndTV, TreeView_GetSelection(m_hwndTV))
  If (tvid Is Nothing) = False Then
    pidlFQ = CopyPIDL(tvid.pidlFQ)
    pidlRel = CopyPIDL(tvid.pidlRel)
    Set tvid = Nothing
  
    ' Clear and reload the ListView.
    Call ClearListView(ListView1)
    Call FillListView(ListView1, pidlFQ, pidlRel)
    
    ' Free the pidls we just copied
    Call FreePIDL(pidlFQ)
    Call FreePIDL(pidlRel)
  End If
  
  Screen.MousePointer = vbDefault

End Sub
