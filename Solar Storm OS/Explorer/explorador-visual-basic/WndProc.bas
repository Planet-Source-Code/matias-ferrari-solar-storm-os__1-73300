Attribute VB_Name = "mWndProc"
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
' A general purpose subclassing module w/ debugging code
'
' - Code was developed using, and is formatted for, 8pt. MS Sans Serif font
' ==============================================

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As Any) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Const WS_BORDER = &H800000
Public Const WS_CLIPCHILDREN = &H2000000

Public Enum GWL_nIndex
  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
End Enum

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long

Private Const OLDWNDPROC = "OldWndProc"
Private Const OBJECTPTR = "ObjectPtr"

' Allocated string pointer filled with OLDWNDPROC string and used with all
' CallWindowProc(GetProp(hwnd, m_lpszOldWndProc), ...) calls.
Private m_lpszOldWndProc As Long

#If DEBUGWINDOWPROC Then
  ' maintains a WindowProcHook reference for each subclassed window.
  ' the window's handle is the collection item's key string.
  Private m_colWPHooks As New Collection
#End If
'

' On first window subclass, allocates a memory buffer and copies the
' "OldWndProc" string to the buffer. On last unsubclass, frees and
' zeros the allocated buffer. The pointer to the buffer is passed directly
' to GetProp when retrieing the subclassed window's original window
' procedure pointer, eliminating VB's Unicode to ANSI string conversion
' in our window procedures.

Private Sub SetWndProcPropertyBuffer(hWnd As Long, fAdd As Boolean)
  Static colhWnds As New Collection
  
  ' Collection holds the handles of all subclassed windows,
  ' ensuring an accurate count of unique handles.
  On Error Resume Next
  If fAdd Then
    colhWnds.Add hWnd, CStr(hWnd)
  Else
    colhWnds.Remove CStr(hWnd)
  End If
  On Error GoTo 0
  
  ' If adding a window handle and the buffer is not yet
  ' allocated, allocate it.
  If fAdd Then
    If (m_lpszOldWndProc = 0) Then
      m_lpszOldWndProc = LocalAlloc(LPTR, Len(OLDWNDPROC))
      If m_lpszOldWndProc Then
        Call lstrcpyA(ByVal m_lpszOldWndProc, ByVal OLDWNDPROC)
'Debug.Print "wndproc buffer allocated"
      End If
    End If
  
  ' If removing a window handle, the collection count is zero, and the
  ' buffer is allocated, deallocate the buffer memory and zero the variable
  ElseIf (fAdd = False) And (colhWnds.Count = 0) Then
    If m_lpszOldWndProc Then
      Call LocalFree(m_lpszOldWndProc)
      m_lpszOldWndProc = 0
'Debug.Print "wndproc buffer freed"
    End If
  End If   ' fAdd

End Sub

Public Function SubClass(hWnd As Long, _
                                         lpfnNew As Long, _
                                         Optional objNotify As Object = Nothing) As Boolean
  Dim lpfnOld As Long
  Dim fSuccess As Boolean
  On Error GoTo Out

  If GetProp(hWnd, OLDWNDPROC) Then
    SubClass = True
    Exit Function
  End If
  
  Call SetWndProcPropertyBuffer(hWnd, True)
  
#If (DEBUGWINDOWPROC = 0) Then
    lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, lpfnNew)

#Else
    Dim objWPHook As WindowProcHook
    
    Set objWPHook = CreateWindowProcHook
    m_colWPHooks.Add objWPHook, CStr(hWnd)
    
    With objWPHook
      Call .SetMainProc(lpfnNew)
      lpfnOld = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
      Call .SetDebugProc(lpfnOld)
    End With

#End If
  
  If lpfnOld Then
    fSuccess = SetProp(hWnd, OLDWNDPROC, lpfnOld)
    If (objNotify Is Nothing) = False Then
      fSuccess = fSuccess And SetProp(hWnd, OBJECTPTR, ObjPtr(objNotify))
    End If
  End If
  
Out:
  If fSuccess Then
    SubClass = True
  
  Else
    If lpfnOld Then Call SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld)
    MsgBox "Error subclassing window &H" & Hex(hWnd) & vbCrLf & vbCrLf & _
                  "Err# " & Err.Number & ": " & Err.Description, vbExclamation
  End If
  
End Function

Public Function UnSubClass(hWnd As Long) As Boolean
  Dim lpfnOld As Long
  
  lpfnOld = GetProp(hWnd, OLDWNDPROC)
  If lpfnOld Then
    
    If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOld) Then
      Call RemoveProp(hWnd, OLDWNDPROC)
      Call RemoveProp(hWnd, OBJECTPTR)

#If DEBUGWINDOWPROC Then
      ' remove the WindowProcHook reference from the collection
      m_colWPHooks.Remove CStr(hWnd)
#End If
      
      Call SetWndProcPropertyBuffer(hWnd, False)
      UnSubClass = True
    
    End If   ' SetWindowLong
  End If   ' lpfnOld

End Function

' Processes Form1.TreeView1 window messages

Public Function TVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg

    ' ============================================================
    ' Prevent the TreeView from removing our system imagelist assignment, which
    ' it wil do when it sees no VB ImageList associated with it.
    ' (the TreeView can't be subclassed when we're assigning imagelists...)
    
    Case TVM_SETIMAGELIST
      Exit Function
      
    ' ============================================================
    ' Process TreeView notification messages reflected back to the TreeView from
    ' it's parent OLE control reflector window.

    Case OCM_NOTIFY
      Dim dwRtn As Long
  
      ' If the notification is non-zero, it's cancelled...
      dwRtn = DoTVNotify(hWnd, lParam)
      If dwRtn Then
        TVWndProc = dwRtn
        Exit Function
      End If
    
    ' ============================================================
    ' Handle owner-draw context menu messages (for the Send To submenu)
    
    Case WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM
    
      If (ICtxMenu2 Is Nothing) = False Then
        Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
      End If
    
    ' ============================================================
    ' Unsubclass the window.
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
  
  End Select
  
  TVWndProc = CallWindowProc(GetProp(hWnd, m_lpszOldWndProc), hWnd, uMsg, wParam, lParam)

End Function

' Processes Form1.ListView1 window messages

Public Function LVWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Select Case uMsg
    
    ' ============================================================
    ' Prevent the ListView from removing our system imagelist assignment, which
    ' it wil do when it sees no VB ImageList associated with it.
    ' (the ListView can't be subclassed when we're assigning imagelists...)
    
    Case LVM_SETIMAGELIST
      Exit Function
    
    ' ============================================================
    ' Process ListView notification messages reflected back to the ListView from
    ' it's parent OLE control reflector window.

    Case OCM_NOTIFY
      Dim dwRtn As Long
  
      ' If the notification is non-zero, it's cancelled...
      dwRtn = DoLVNotify(hWnd, lParam)
      If dwRtn Then
        LVWndProc = dwRtn
        Exit Function
      End If
    
    ' ============================================================
    ' Handle owner-draw context menu messages (for the Send To submenu)
    
    Case WM_INITMENUPOPUP, WM_DRAWITEM, WM_MEASUREITEM
    
      If (ICtxMenu2 Is Nothing) = False Then
        Call ICtxMenu2.HandleMenuMsg(uMsg, wParam, lParam)
      End If
    
    ' ======================================================
    ' Unsubclass the window.
    
    Case WM_DESTROY
      ' OLDWNDPROC will be gone after UnSubClass is called!
      Call CallWindowProc(GetProp(hWnd, OLDWNDPROC), hWnd, uMsg, wParam, lParam)
      Call UnSubClass(hWnd)
      Exit Function
      
  End Select
  
  LVWndProc = CallWindowProc(GetProp(hWnd, m_lpszOldWndProc), hWnd, uMsg, wParam, lParam)
  
End Function
