VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVHeaderSortIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' We're using IE3 definitions
#Const WIN32_IE = &H300

Private m_ListView As ListView
Attribute m_ListView.VB_VarHelpID = -1
Private m_himl As Long

Public Enum SortOrderConstants
  soAscending = 0     ' lvwAscending
  soDescending = 1  ' lvwDescending
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long

' ============================================================
' listview definitions

' messages
Private Const LVM_FIRST = &H1000
#If (WIN32_IE >= &H300) Then
Private Const LVM_GETHEADER = (LVM_FIRST + 31)
#End If

' ============================================================
' header definitions

' messages
Private Const HDM_FIRST = &H1200
Private Const HDM_SETITEM = (HDM_FIRST + 4)
#If (WIN32_IE >= &H300) Then
Private Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
#End If

Private Type HDITEM   ' was HD_ITEM
  mask As Long
  cxy As Long
  pszText As String   ' if retrieving text, must be pre-allocated
  hbm As Long
  cchTextMax As Long
  fmt As Long
  lParam As Long
#If (WIN32_IE >= &H300) Then
  iImage As Long       ' index of bitmap in ImageList
  iOrder As Long        ' where to draw this item
#End If
End Type

' HDITEM mask
Private Const HDI_FORMAT = &H4
#If (WIN32_IE >= &H300) Then
Private Const HDI_IMAGE = &H20
#End If

' HDITEM fmt
Private Const HDF_LEFT = 0
Private Const HDF_RIGHT = 1
#If (WIN32_IE >= &H300) Then
Private Const HDF_IMAGE = &H800
Private Const HDF_BITMAP_ON_RIGHT = &H1000
#End If
Private Const HDF_STRING = &H4000

' ============================================================
' imagelist definitions

' ImageList_Create flags
Private Const ILC_MASK = &H1
Private Const ILC_COLOR8 = &H8

Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long
'

Private Sub Class_Initialize()
  m_himl = ImageList_Create(16, 16, ILC_MASK Or ILC_COLOR8, 2, 0)
  If m_himl Then
    ' Load the icons into the image so that their zero-based
    ' indicescorrespond to the SortOrderConstants values.

  End If
End Sub

Private Sub Class_Terminate()
  If m_himl Then Call ImageList_Destroy(m_himl)
End Sub

Public Property Get ListView() As ListView
  Set ListView = m_ListView
End Property

Public Property Set ListView(lv As Object)
  Set m_ListView = lv
End Property

' Sets and removes header sort order icons

Public Function SetHeaderIcons(iActiveColumn As Long, iSortOrder As SortOrderConstants) As Boolean
  Static hwndHdr As Long
  Dim i As Long
  Dim fShow As Boolean
  Dim fAlignRight As Boolean
  Dim hdi As HDITEM
  
  If (m_himl = 0) Or (m_ListView Is Nothing) Then Exit Function
  If (m_ListView.View <> lvwReport) Then Exit Function
  
  ' The ListView's header is created *after* the first
  ' ColumnHeader is added.
  If (hwndHdr = 0) Then
    hwndHdr = ListView_GetHeader(m_ListView.hWnd)
    Call Header_SetImageList(hwndHdr, m_himl)
  End If
  
  If (hwndHdr = 0) Then Exit Function
    
  With m_ListView.ColumnHeaders
    For i = 0 To .Count - 1
      hdi.mask = HDI_FORMAT Or HDI_IMAGE
      
      fAlignRight = .Item(i + 1).Alignment = lvwColumnRight
      
      ' Since we're setting the header's format, we have to
      ' specify the string flag
      hdi.fmt = HDF_STRING Or (fAlignRight And HDF_RIGHT)   ' HDF_LEFT = 0
      
      ' If the active column, add the sort icon with the appropriate
      ' alignment (the icon is removed if HDF_IMAGE is not set).
      If (i = iActiveColumn) Then
        hdi.fmt = hdi.fmt Or HDF_IMAGE Or ((fAlignRight = False) And HDF_BITMAP_ON_RIGHT)
      End If
      
      ' If not the soAscending icon index (0), then set to
      ' the soDescending (1) icon index.
      hdi.iImage = Abs(CBool(iSortOrder))
      
      Call Header_SetItem(hwndHdr, i, hdi)
    Next
  End With
  
  SetHeaderIcons = True
  
End Function

' ============================================================
' listview, header, imagelist macros

#If (WIN32_IE >= &H300) Then

Private Function ListView_GetHeader(hWnd As Long) As Long
  ListView_GetHeader = SendMessage(hWnd, LVM_GETHEADER, 0, 0)
End Function
'
#End If
'

Private Function Header_SetItem(hwndHD As Long, i As Long, phdi As HDITEM) As Boolean
  Header_SetItem = SendMessage(hwndHD, HDM_SETITEM, ByVal i, phdi)
End Function
 
#If (WIN32_IE >= &H300) Then

Private Function Header_SetImageList(hWnd As Long, himl As Long) As Long
  Header_SetImageList = SendMessage(hWnd, HDM_SETIMAGELIST, 0, ByVal himl)
End Function
'
#End If
'

Private Function ImageList_AddIcon(himl As Long, hIcon As Long) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function
