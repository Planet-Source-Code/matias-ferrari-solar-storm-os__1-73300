Attribute VB_Name = "Buscar_unidades"

Option Explicit

' Api Declarations
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Drive type constants
Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6


Public Enum DriveTypeConst
    [REMOVABLE] = 2
    [Fixed] = 3
    [REMOTE] = 4
    [CDROM] = 5
    [RAMDISK] = 6
End Enum
Public Function GetDrives(ByRef LstBox As ListBox, Optional DriveModel As DriveTypeConst) As Integer
    Dim ret As Long, AllDrives As String, IsolatedDrive As String, Posn As Long, DriveType As Long
    Dim NumDrives As Integer
   ' LstBox.Clear
    AllDrives = Space(64)
    ret = GetLogicalDriveStrings(Len(AllDrives), AllDrives)
    AllDrives = Left(AllDrives, ret)
    Do
    Posn = InStr(AllDrives, Chr(0))
    If Posn Then
            IsolatedDrive = Left(AllDrives, Posn)
            AllDrives = Mid(AllDrives, Posn + 1, Len(AllDrives))
            DriveType& = GetDriveType(IsolatedDrive)
            If DriveModel = 0 Then
                  LstBox.AddItem Mid(IsolatedDrive, 1, 2)
                  NumDrives = NumDrives + 1
            Else
                If DriveType = DriveModel Then
                    LstBox.AddItem Mid(IsolatedDrive, 1, 2)
                    NumDrives = NumDrives + 1
                End If
            End If
            
          End If
      Loop Until AllDrives = ""
      GetDrives = NumDrives
End Function




