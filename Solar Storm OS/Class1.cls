VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Constantes que indican la sección y clave donde se almacenarán _
 los accesos directos en el registro de windows

Public CLAVE As String
Public SECCION As String


' Api GetOpenFileName  - Permite abrir el cuadro _
  de diálogo para seleccionar un archivo

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

' Estructure OPENFILENAME que requiere el Api GetOpenFileName

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' Colección para almacenar los accesos directos
Private Col_Ruta As Collection

'Cantidad de accesos que se van a utilizar ( copia local )
Private m_Count As Byte

'Variables locales para hacer referencia a los menues _
 de los accesos directos y para el separador
Dim m_El_Menu As Object
Dim m_Menu_Separador As Object

'Inicia la clase. Recibe como parámetro el nombre del menú _
 para los accesos directos y el nombre del menú para el separador
 
Sub Init(MenuReciente As Object, MenuSeparador As Object)

Dim i As Integer
Dim s_Ruta As String
    
    'Nueva colecciónque mantiene las rutas de los ficheros
    Set Col_Ruta = New Collection
    
    MenuReciente(1).Visible = False
    
    'Crea menues dinámicos
    For i = 2 To 25
        Load MenuReciente(i)
        MenuReciente(i).Visible = False
    Next
    
    'Lee del registro todas las entradas
    For i = 1 To m_Count
        
        s_Ruta = GetSetting(CLAVE, SECCION, CStr(i), vbNullString)
        
        'si no hay entradas creadas sale
        If Len(s_Ruta) = 0 Then
            Exit For
        End If
        
        'Añade a la colección
        Col_Ruta.Add s_Ruta
    
    Next i
    
    'Asigna el menú a la variable local para poder _
     utilizarla dentro del módulo
    If m_El_Menu Is Nothing Then
       Set m_El_Menu = MenuReciente
    End If
    
    If m_Menu_Separador Is Nothing Then
       Set m_Menu_Separador = MenuSeparador
    End If
    
    ' Carga los accesos en el menú (si es que hay) leyendo del registro
    Call MostrarLista
    
End Sub

' Extrae solo el nombre del archivo de la ruta
Private Function Obtener_Nombre(s_Path As String)

Obtener_Nombre = Right(s_Path, Len(s_Path) - InStrRev(s_Path, "\"))

End Function

'Añade un nuevo acceso directo al menú y también _
 graba los datos en el registro
Sub NuevoElemento(ByVal Ruta_Archivo As String)

Dim i As Integer

    ' ... si es el primero
    If Col_Ruta.Count = 0 Then
        Col_Ruta.Add Ruta_Archivo
    Else
        Col_Ruta.Add Ruta_Archivo, , 1
    End If
    
    'Elimina el dato de la colección si está repetido
    For i = Col_Ruta.Count To 2 Step -1
        If Col_Ruta(i) = Ruta_Archivo Then
            Col_Ruta.Remove i
        End If
    Next i
    
    
    Do While Col_Ruta.Count > m_Count
        Col_Ruta.Remove m_Count + 1
    Loop

    'Carga los datos en los menues y graba las entradas
    Call MostrarLista
    Call GrabarLista
    
End Sub


Private Sub MostrarLista()
Dim i As Integer
    
    'Recorre la colección y le coloca a cada menú el _
     caption (nombre de archivo) y también lo hace visible
    
    For i = 1 To Col_Ruta.Count
        m_El_Menu(i).Caption = CStr(i) & " " & Obtener_Nombre(Col_Ruta(i))
        m_El_Menu(i).Visible = True
    Next i
    
    ' hace visible el menú separador = al primer elemento
    m_Menu_Separador.Visible = m_El_Menu(1).Visible
    
End Sub

'graba en el registro las entradas para los accesos directos

Private Sub GrabarLista()
Dim i As Integer
    
    'Recorre la colección
    For i = 1 To Col_Ruta.Count
        
        Call SaveSetting(CLAVE, SECCION, CStr(i), Col_Ruta(i))
    
    Next i

    On Local Error Resume Next
    
    For i = Col_Ruta.Count + 1 To m_Count
        DeleteSetting CLAVE, SECCION, CStr(i)
    Next i

End Sub


' Sub que recorre elimina la lista de accesos directos, _
  tanto del menú como del registro de windows

Sub EliminarLista()

    Dim i As Integer

    On Local Error Resume Next
    
    'elimina la sección completa del registro
    Call DeleteSetting(CLAVE, SECCION)
    
    ' Oculta todos los menues
    For i = 1 To m_El_Menu.Count
        m_El_Menu(i).Visible = False
    Next
    
    Set Col_Ruta = Nothing
    Set Col_Ruta = New Collection
    
    'Oculta el menú separador
    m_Menu_Separador.Visible = False
    
End Sub

'Función que retorna del Path del archivo, solo el nombre del mismo
Function ObtenerPath(indice_Menu As Integer) As String

ObtenerPath = Col_Ruta(indice_Menu)

End Function

Function CommonDialog_Abrir(Hwnd_form As Long, _
                            Optional Filtro As String, _
                            Optional path_Inicial As String, _
                            Optional Dialog_title As String) As String
                            

' Para la estructura del cuadro de diálogo
Dim T_Open_File As OPENFILENAME
    
    
    'Filtro por defecto si no se pasa como parámetro
    If Filtro = vbNullString Then
       Filtro = "Todos los archivos(*.*)" + Chr$(0) + "*.*" + Chr$(0)
    End If
    
    If Dialog_title = vbNullString Then
       Dialog_title = " Abrir archivo ... "
    End If
    
    T_Open_File.lStructSize = Len(T_Open_File)
    'Establecemos el hwnd del form
    T_Open_File.hwndOwner = Hwnd_form
    'Establecemos la instancia de la aplicacion
    T_Open_File.hInstance = App.hInstance
    'filtro de extensiones de archivos a mostrar
    T_Open_File.lpstrFilter = Filtro
    'Acá debemos indicar un buffer para el nombre de archivo
    T_Open_File.lpstrFile = Space$(254)
    
    T_Open_File.nMaxFile = 255
    'Buffer para el titulo del archivo
    T_Open_File.lpstrFileTitle = Space$(254)
    T_Open_File.nMaxFileTitle = 255
    
    'Establecemos el directorio por defecto que mostrará el dialogo al abrir
    T_Open_File.lpstrInitialDir = path_Inicial
    
    ' título del cuadro de diálogo
    T_Open_File.lpstrTitle = Dialog_title
    
    'Flags para el Common dialog
    T_Open_File.flags = 0

    'Abrimos el cuadro de diálogo pasandole a GetOpenFileName la estructura anterior
    If GetOpenFileName(T_Open_File) Then
        
        CommonDialog_Abrir = Trim$(Replace(T_Open_File.lpstrFile, _
                                           Chr(0), vbNullString))
    End If
    
End Function

'Propiedad que mantiene la cantidad _
 de accesos directos a utilizar
Public Property Get Count() As Byte
    Count = m_Count
End Property

Public Property Let Count(Cantidad As Byte)
On Local Error GoTo errProp
    
    m_Count = Cantidad

Exit Property

errProp:

MsgBox Err.Number & " : " & Err.Description, vbCritical

End Property

 



