Attribute VB_Name = "Buscar"
Option Explicit


'***************************************************************************
'*  Código fuente del módulo bas
'***************************************************************************



'Declaraciones del Api
'------------------------------------------------------------------------------

'Esta función busca el primer archivo de un Dir
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" ( _
    ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long

'Esta el siguiente archivo o directorio
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" ( _
    ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" ( _
    ByVal lpFileName As String) As Long

'Esta cierra el Handle de búsqueda
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


' Constantes
'------------------------------------------------------------------------------

'Constantes de atributos de archivos
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

'Otras constantes
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1


'UDT
'------------------------------------------------------------------------------

'Estructura para las fechas de los archivos
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Estructura necesaria para la información de archivos
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


'-----------------------------------------------------------------------
    'Funciones
'-----------------------------------------------------------------------


'Esta función es para formatear los nombres de archivos y directorios. Elimina los CHR(0)
'------------------------------------------------------------------------
Function Eliminar_Nulos(OriginalStr As String) As String
    
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    Eliminar_Nulos = OriginalStr

End Function

'Esta función es la principal que permite buscar _
 los archivos y listarlos en el ListBox


Function FindFilesAPI(Path As String, _
                      SearchStr As String, _
                      FileCount As Long, _
                      DirCount As Long, _
                      ListBox As ListBox)


    Dim FileName As String
    Dim DirName As String
    Dim dirNames() As String
    Dim nDir As Long
    Dim i As Long
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Long


    If Right(Path, 1) <> "\" Then Path = Path & "\"
        ' Buscamos por mas directorios
        nDir = 0
        ReDim dirNames(nDir)
        Cont = True
        hSearch = FindFirstFile(Path & "*", WFD)
            If hSearch <> INVALID_HANDLE_VALUE Then
                Do While Cont
                    DirName = Eliminar_Nulos(WFD.cFileName)
                    ' Ignore the current and encompassing directories.
                    If (DirName <> ".") And (DirName <> "..") Then
                        ' Check for directory with bitwise comparison.
                            If GetFileAttributes(Path & DirName) _
                                And FILE_ATTRIBUTE_DIRECTORY Then
                                
                                dirNames(nDir) = DirName
                                DirCount = DirCount + 1
                                nDir = nDir + 1
                                ReDim Preserve dirNames(nDir)
                            
                            End If
                    End If
                    Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
                Loop
                
                Cont = FindClose(hSearch)
            
            End If

        hSearch = FindFirstFile(Path & SearchStr, WFD)
        Cont = True
        If hSearch <> INVALID_HANDLE_VALUE Then
            While Cont
                FileName = Eliminar_Nulos(WFD.cFileName)
                    If (FileName <> ".") And (FileName <> "..") Then
                        FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) _
                                                                  + WFD.nFileSizeLow
                        FileCount = FileCount + 1
                        ListBox.AddItem Path & FileName
                    End If
                Cont = FindNextFile(hSearch, WFD) ' Get next file
            Wend
        Cont = FindClose(hSearch)
        End If

        ' Si estos son Sub Directorios......
        If nDir > 0 Then

        For i = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", _
                                                SearchStr, FileCount, DirCount, ListBox)
        Next i
    End If
End Function


