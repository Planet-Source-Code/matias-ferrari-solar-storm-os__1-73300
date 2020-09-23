Attribute VB_Name = "Module1"
Public FrmDoc As frmDocumento
Public nForms As Integer
Public position As Integer

Public FrmFocus As Form

Public Const sCaption = "Nuevo documento sin título "

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Option Explicit

Public Declare Function SetErrorMode _
    Lib "kernel32" ( _
    ByVal wMode As Long) As Long

Public Declare Sub InitCommonControls Lib "Comctl32" ()

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
                ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function InvalidateRect Lib "user32" _
                (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Enum enuTBType
    enuTB_FLAT = 1
    enuTB_STANDARD = 2
End Enum

Private Const GCL_HBRBACKGROUND = (-10)

Public Sub CambiarFondoToolbar(TB As Object, PNewBack As Long, pType As enuTBType)
Dim lTBWnd      As Long

    Select Case pType
        
        Case enuTB_FLAT
            DeleteObject SetClassLong(TB.hwnd, GCL_HBRBACKGROUND, PNewBack)
        
        Case enuTB_STANDARD
            lTBWnd = FindWindowEx(TB.hwnd, 0, "msvb_lib_toolbar", vbNullString)
            DeleteObject SetClassLong(lTBWnd, GCL_HBRBACKGROUND, PNewBack)
    End Select
End Sub

Sub TopMost(Frm As Form)
    SetWindowPos Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
