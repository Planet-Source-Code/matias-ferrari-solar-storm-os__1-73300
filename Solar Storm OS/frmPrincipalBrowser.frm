VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPrincipalBrowser 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrincipalBrowser.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPrincipalBrowser.frx":0CCA
   ScaleHeight     =   8010
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   285
      Left            =   6480
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlPrincipal_bn 
      Left            =   8880
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":18CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":1E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":2400
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":299A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":2F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":34CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":3A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":4002
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":459C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":4B36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPrincipal 
      Left            =   8280
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":50D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":566A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":5C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":619E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":6738
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":6CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":726C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":7806
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":7DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":833A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":88D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":8E6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":9408
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":99A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":9F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":A4D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":AA70
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":B00A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipalBrowser.frx":B5A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pbxSeparador 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   11610
      TabIndex        =   7
      Top             =   975
      Width           =   11610
   End
   Begin VB.PictureBox pbxContenedor 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6705
      Left            =   0
      ScaleHeight     =   6705
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   1005
      Width           =   6000
      Begin SHDocVwCtl.WebBrowser wbrPrincipal 
         Height          =   5415
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
         ExtentX         =   9975
         ExtentY         =   9551
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSComctlLib.TabStrip TS 
         Height          =   5775
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   10186
         MultiRow        =   -1  'True
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         ImageList       =   "imlPrincipal"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
               ImageIndex      =   18
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox cbrPrincipal 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11610
      TabIndex        =   1
      Top             =   0
      Width           =   11610
      Begin MSComctlLib.Toolbar tbrPrincipal 
         Height          =   330
         Left            =   10440
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlPrincipal_bn"
         HotImageList    =   "imlPrincipal"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nue"
               Object.ToolTipText     =   "Nuevo explorador"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "atr"
               Object.ToolTipText     =   "Atrás"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ade"
               Object.ToolTipText     =   "Adelante"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "det"
               Object.ToolTipText     =   "Detener"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "act"
               Object.ToolTipText     =   "Actualizar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ini"
               Object.ToolTipText     =   "Inicio"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bus"
               Object.ToolTipText     =   "Búsqueda"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "fav"
               Object.ToolTipText     =   "Favoritos"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "his"
               Object.ToolTipText     =   "Historial"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "imp"
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboURL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   200
         Width           =   8325
      End
      Begin VB.Image BtnCerrar_Pestaña 
         Height          =   300
         Left            =   3240
         Picture         =   "frmPrincipalBrowser.frx":BB3E
         ToolTipText     =   "Cerrar Pestaña"
         Top             =   600
         Width           =   330
      End
      Begin VB.Image BtnVista_Preliminar 
         Height          =   270
         Left            =   1920
         Picture         =   "frmPrincipalBrowser.frx":C0D0
         ToolTipText     =   "Vista Preliminar"
         Top             =   600
         Width           =   270
      End
      Begin VB.Image BtnCodigo_Fuente 
         Height          =   270
         Left            =   5160
         Picture         =   "frmPrincipalBrowser.frx":C502
         ToolTipText     =   "Ver Codigo Fuente"
         Top             =   600
         Width           =   270
      End
      Begin VB.Image BtnSave 
         Height          =   270
         Left            =   1560
         Picture         =   "frmPrincipalBrowser.frx":CAAC
         ToolTipText     =   "Guardar Pagina"
         Top             =   600
         Width           =   270
      End
      Begin VB.Image BtnOpen 
         Height          =   285
         Left            =   1200
         Picture         =   "frmPrincipalBrowser.frx":CEDE
         ToolTipText     =   "Abrir Pagina Local"
         Top             =   600
         Width           =   300
      End
      Begin VB.Image BtnNueva_Ventana 
         Height          =   330
         Left            =   2880
         Picture         =   "frmPrincipalBrowser.frx":D394
         ToolTipText     =   "Nueva Ventana"
         Top             =   600
         Width           =   330
      End
      Begin VB.Image BtnNueva_pestaña 
         Height          =   300
         Left            =   2520
         Picture         =   "frmPrincipalBrowser.frx":D9AE
         ToolTipText     =   "Nueva Pestaña"
         Top             =   600
         Width           =   300
      End
      Begin VB.Image BtnPrint 
         Height          =   300
         Left            =   5520
         Picture         =   "frmPrincipalBrowser.frx":DEA0
         ToolTipText     =   "Imprimir"
         Top             =   600
         Width           =   300
      End
      Begin VB.Image btnSearch 
         Height          =   300
         Left            =   4560
         Picture         =   "frmPrincipalBrowser.frx":E392
         ToolTipText     =   "Buscar..."
         Top             =   600
         Width           =   300
      End
      Begin VB.Image BtnFavoritos 
         Height          =   330
         Left            =   4200
         Picture         =   "frmPrincipalBrowser.frx":E884
         ToolTipText     =   "Añadir a Favoritos"
         Top             =   600
         Width           =   330
      End
      Begin VB.Image btnHome 
         Height          =   300
         Left            =   3840
         Picture         =   "frmPrincipalBrowser.frx":EE9E
         ToolTipText     =   "Pagina Principal"
         Top             =   600
         Width           =   285
      End
      Begin VB.Image BtnDetener 
         Height          =   315
         Left            =   9960
         Picture         =   "frmPrincipalBrowser.frx":F390
         ToolTipText     =   "Detener"
         Top             =   240
         Width           =   315
      End
      Begin VB.Image BtnRefrescar 
         Height          =   300
         Left            =   9600
         Picture         =   "frmPrincipalBrowser.frx":F912
         ToolTipText     =   "Refrescar"
         Top             =   240
         Width           =   300
      End
      Begin VB.Image btnBack 
         Height          =   450
         Left            =   120
         Picture         =   "frmPrincipalBrowser.frx":FE04
         ToolTipText     =   "Atras"
         Top             =   120
         Width           =   465
      End
      Begin VB.Image BtAdelante 
         Height          =   450
         Left            =   600
         Picture         =   "frmPrincipalBrowser.frx":10986
         ToolTipText     =   "Adelante"
         Top             =   120
         Width           =   465
      End
   End
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7710
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoNuevo 
         Caption         =   "&Nueva pestaña"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuNuevaVentana 
         Caption         =   "&Nueva ventana"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuCerrarPestaña2 
         Caption         =   "Cerrar pestaña"
      End
      Begin VB.Menu mnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrirPagina 
         Caption         =   "Abrir página local"
      End
      Begin VB.Menu mnuArchivoGuardarComo 
         Caption         =   "Guardar Página web..."
      End
      Begin VB.Menu mnuArchivoSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivoConfigurarPagina 
         Caption         =   "Con&figurar página..."
      End
      Begin VB.Menu mnuArchivoImprimir 
         Caption         =   "&Imprimir..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuArchivoVistaPreliminar 
         Caption         =   "Vista p&reliminar..."
      End
      Begin VB.Menu mnuArchivoSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivoPropiedades 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReciente 
         Caption         =   "aaaaaaaaa"
         Index           =   1
      End
      Begin VB.Menu nulineadsfdsf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivoSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuEdicionCortar 
         Caption         =   "C&ortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdicionCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdicionPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdicionSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicionSeleccionarTodo 
         Caption         =   "&Seleccionar todo"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEdicionSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicionBuscar 
         Caption         =   "&Buscar en esta página..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuPaginaInicio 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuPageInicio 
         Caption         =   "&Establecer página de inicio"
      End
      Begin VB.Menu mnulineaOpcione1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEliminarURLVistadas 
         Caption         =   "Eliminar Listas de url visitadas"
      End
      Begin VB.Menu mnuLinea4 
         Caption         =   "-"
      End
      Begin VB.Menu sadasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerCodigoFuente 
         Caption         =   "Ver código fuente"
      End
   End
   Begin VB.Menu mnuFavoritos 
      Caption         =   "&Favoritos"
      Begin VB.Menu mnuAgregarFavoritos 
         Caption         =   "Añadir esta página"
      End
      Begin VB.Menu mnuSeparadorFavoritos 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUrlFavoritos 
         Caption         =   "aaa"
         Index           =   1
      End
      Begin VB.Menu fdgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEliminarFavoritos 
         Caption         =   "Eliminar toda la lista"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "&Acerca de vbNavegador..."
      End
   End
   Begin VB.Menu mPopup 
      Caption         =   "menuContext"
      Visible         =   0   'False
      Begin VB.Menu mCerrarTab 
         Caption         =   "Cerrar pestaña"
      End
      Begin VB.Menu mnuCerrarTodo 
         Caption         =   "Cerrar todas las pestañas"
      End
      Begin VB.Menu mnuCerrarPestañaNoActual 
         Caption         =   "&Cerrar todas menos la actual"
      End
      Begin VB.Menu mnulin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnupopRefresh 
         Caption         =   "&Refrescar"
      End
      Begin VB.Menu mnuRefrescarTodo 
         Caption         =   "Refrescar Todas las pestañas"
      End
      Begin VB.Menu linearefrescar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerCodigoFuenteP2 
         Caption         =   "Ver código fuente"
      End
      Begin VB.Menu dasdasd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavoritos2 
         Caption         =   "Añadir a favoritos"
      End
   End
End
Attribute VB_Name = "frmPrincipalBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" ( _
    ByVal hwnd As Long, _
    ByVal szApp As String, _
    ByVal szOtherStuff As String, _
    ByVal hIcon As Long) As Long

Dim Pagina As String
Dim c_Recientes As Class1
Dim cFavoritos As Class1

Property Get Wbr() As SHDocVwCtl.WebBrowser
    On Error Resume Next
    'Devuelve el control «WebBrowser» activo (el de la pestaña seleccionada)
    If Not TS.SelectedItem Is Nothing Then
        Set Wbr = wbrPrincipal(TS.SelectedItem.Tag)
            Wbr.Silent = True
    End If
End Property

Private Sub BtAdelante_Click()
On Error Resume Next
                Wbr.GoForward
            
End Sub

Private Sub btnBack_Click()
 On Error Resume Next
            Wbr.GoBack
                
End Sub

Private Sub BtnCerrar_Pestaña_Click()
mnuCerrarPestaña2_Click
End Sub

Private Sub BtnCodigo_Fuente_Click()
mnuVerCodigofuente_Click
End Sub

Private Sub BtnDetener_Click()
Wbr.Stop

End Sub

Private Sub BtnFavoritos_Click()
mnuAgregarFavoritos_Click
End Sub

Private Sub btnHome_Click()
 Wbr.Navigate Pagina
End Sub

Private Sub BtnNueva_pestaña_Click()
mnuArchivoNuevo_Click
End Sub

Private Sub BtnNueva_Ventana_Click()
mnuNuevaVentana_Click
End Sub

Private Sub BtnOpen_Click()
mnuAbrirPagina_Click
End Sub

Private Sub BtnPrint_Click()
 On Error Resume Next
            mnuArchivoImprimir_Click
End Sub

Private Sub BtnRefrescar_Click()
RefrescarPagina
End Sub

Private Sub BtnSave_Click()
mnuArchivoGuardarComo_Click
End Sub

Private Sub btnSearch_Click()
Wbr.GoSearch
End Sub

Private Sub BtnVista_Preliminar_Click()
mnuArchivoVistaPreliminar_Click
End Sub

Private Sub cboURL_Change()
cboURL.BackColor = &HC0FFFF
cboURL.Font.Bold = True
End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)
    'Al pulsar la tecla {Enter} navegar a la URL introducida (si no está en blanco)
    On Error Resume Next

    If KeyAscii = vbKeyReturn Then
        If Len(cboURL.Text) Then
            If Not TS.SelectedItem Is Nothing Then
                Wbr.Navigate cboURL.Text
                AgregarURL cboURL.Text
                AgregarAccesoDirecto cboURL.Text
            Else
                AbrirNuevo cboURL.Text
                
            End If
        End If
    End If
End Sub

Private Sub cboURL_Click()
    Wbr.Navigate cboURL.Text
    AgregarAccesoDirecto cboURL.Text
End Sub

Private Sub AgregarURL(ByVal strURL As String)
Dim i As Long
    'Agregar una dirección a «cboURL» (si no ha sido agregada anteriormente)
    For i = 0 To cboURL.ListCount - 1
        If cboURL.List(i) = strURL Then Exit Sub
    Next i
    cboURL.AddItem strURL
End Sub















Private Sub Form_Load()
         
    Set c_Recientes = New Class1
    
    With c_Recientes
        .SECCION = "SeccionListado"
        .CLAVE = "ListaREcientes"
        .Count = 20
        Call .Init(mnuReciente, mnuSeparador)
        
    End With
         
    Set cFavoritos = New Class1
    
    With cFavoritos
        .SECCION = "Favoritos"
        .CLAVE = "ListaFavoritos"
        .Count = 100
        Call .Init(mnuUrlFavoritos, mnuSeparadorFavoritos)
    End With
         
    Pagina = GetSetting(App.EXEName, "Opciones", "PaginaInicio", "www.google.com")
    
    Load wbrPrincipal(1)
    wbrPrincipal(1).Visible = True
    TS.SelectedItem.Tag = 1
    TS.Tabs(1).Selected = True
    Wbr.Navigate Pagina
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        pbxContenedor.Width = ScaleWidth
        pb.Move (Me.ScaleWidth - pb.Width + 10), Me.ScaleHeight - (pb.Height)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

  If TS.Tabs.Count = 1 Then Unload Me
  
  If TS.Tabs.Count > 1 Then
  
    If MsgBox(" ¿ Salir y cerrar las " & TS.Tabs.Count & " pestañas abiertas ?", vbQuestion + vbYesNo) = vbYes Then
        Set frmPrincipalBrowser = Nothing
        Unload Me
    Else
        Cancel = True
    End If
  End If

  
End Sub









Private Sub mCerrarTab_Click()
    Unload wbrPrincipal(TS.SelectedItem.Tag)
    TS.Tabs.Remove (TS.SelectedItem.index)
    If TS.Tabs.Count > 0 Then
        TS.Tabs(1).Selected = True
    End If
    
End Sub

Private Sub mnuAbrirPagina_Click()
    Dim Archivo As String
    Archivo = c_Recientes.CommonDialog_Abrir(Me.hwnd, "*.*", "Abrir página")
    
    Wbr.Navigate Archivo
End Sub

Private Sub mnuAgregarFavoritos_Click()
    If Not Wbr Is Nothing Then
        If Wbr.LocationURL <> "" Then
            cFavoritos.NuevoElemento Wbr.LocationURL
            mnuEliminarFavoritos.Enabled = True
        End If
    End If
End Sub

Private Sub mnuArchivoConfigurarPagina_Click()
    Wbr.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuArchivoNuevo_Click()
    AbrirNuevo
End Sub

Sub AbrirNuevo(Optional DirURL As String)
    'Agregar un «WebBrowser» y una pestaña y seleccionarla
    
    If DirURL <> "" Then
        Load wbrPrincipal(wbrPrincipal.UBound + 1)
        With wbrPrincipal(wbrPrincipal.UBound)
            .Move 75, 375, pbxContenedor.ScaleWidth - 165, pbxContenedor.ScaleHeight - 450
            .Navigate DirURL
        End With
        TS.Tabs.Add
        TS.Tabs(TS.Tabs.Count).Tag = wbrPrincipal.UBound
        TS.Tabs(TS.Tabs.Count).Selected = True
        TS.Tabs(TS.Tabs.Count).Image = 14
    Else
        Load wbrPrincipal(wbrPrincipal.UBound + 1)
        With wbrPrincipal(wbrPrincipal.UBound)
            .Move 75, 375, pbxContenedor.ScaleWidth - 165, pbxContenedor.ScaleHeight - 450
            .Navigate Pagina
        End With
        TS.Tabs.Add
        TS.Tabs(TS.Tabs.Count).Tag = wbrPrincipal.UBound
        TS.Tabs(TS.Tabs.Count).Selected = True
        TS.Tabs(TS.Tabs.Count).Image = 14
    End If
End Sub

Private Sub mnuArchivoGuardarComo_Click()
    Wbr.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuArchivoImprimir_Click()
    Wbr.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuArchivoPropiedades_Click()
    Wbr.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuArchivoVistaPreliminar_Click()
    Wbr.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuAyudaAcercaDe_Click()
    ShellAbout hwnd, App.Title, "Autor: Rubén Vigón (vigon@mvps.org)", Icon.Handle
End Sub

Private Sub mnuCerrarPestaña2_Click()
    mCerrarTab_Click
End Sub

Private Sub mnuCerrarPestañaNoActual_Click()
    Dim i As Integer
    Dim TabActual As Integer
 
    On Local Error Resume Next
    For i = 1 To wbrPrincipal.Count - 1
        
        If TS.SelectedItem.Tag <> i Then
            Unload wbrPrincipal(i)
        End If
    Next

    
    TabActual = TS.SelectedItem.Tag

    For i = 1 To TS.Tabs.Count
        If TabActual <> TS.Tabs(1).Tag Then
            TS.Tabs.Remove 1
        ElseIf TabActual <> TS.Tabs(TS.Tabs.Count).Tag Then
            TS.Tabs.Remove TS.Tabs.Count
        End If
        
        'If i <> TS.SelectedItem.Tag Then
        '   Unload wbrPrincipal(i)
        'End If
    Next
    
    TS.SelectedItem.Selected = True
    
    On Error GoTo 0
End Sub

Private Sub mnuCerrarTodo_Click()
    Dim i As Integer
    
    If MsgBox("¿¿ Cerrar todos los Tabs ? ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    For i = 1 To TS.Tabs.Count
        Unload wbrPrincipal(TS.Tabs(i).Tag)
    Next
    TS.Tabs.Clear
    On Error GoTo 0
    
End Sub

Private Sub mnuEdicionBuscar_Click()
    On Local Error Resume Next
    Wbr.SetFocus
    SendKeys ("^f")
End Sub

Private Sub mnuEdicionCortar_Click()
    Wbr.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT 'Edición -> Cortar (Ctrl+X)
End Sub

Private Sub mnuEdicionCopiar_Click()
    Wbr.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT 'Edición -> Copiar (Ctrl+C)
End Sub

Private Sub mnuEdicionPegar_Click()
    Wbr.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT 'Edición -> Pegar (Ctrl+V)
End Sub

Private Sub mnuEdicionSeleccionarTodo_Click()
    On Local Error Resume Next
        Wbr.SetFocus
        Wbr.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub mnuEliminarFavoritos_Click()
    If MsgBox("¿ Eliminar la lista ?", vbQuestion + vbYesNo) = vbYes Then
        cFavoritos.EliminarLista
        mnuEliminarFavoritos.Enabled = False
    End If
End Sub

Private Sub mnuEliminarURLVistadas_Click()
    c_Recientes.EliminarLista
End Sub



Private Sub mnuFavoritos2_Click()
    mnuAgregarFavoritos_Click
End Sub

Private Sub mnuNuevaVentana_Click()
    Dim frm_web As frmPrincipalBrowser
    'Nueva instancia del formulario
    Set frm_web = New frmPrincipalBrowser
        
        frm_web.Show
    
End Sub

Private Sub mnuPageInicio_Click()
    
    If MsgBox("¿ Guardar la página página actual como página de inicio ?", vbQuestion + vbYesNo) = vbNo Then
       Exit Sub
    End If
    If cboURL.Text <> "" Then
        Pagina = cboURL.Text
        SaveSetting App.EXEName, "Opciones", "PaginaInicio", Pagina
    End If
End Sub

Private Sub mnupopRefresh_Click()
    RefrescarPagina
End Sub

Sub RefrescarPagina()
    If Not Wbr Is Nothing Then
        Wbr.Refresh
    End If
End Sub

Private Sub mnuRefrescarTodo_Click()
    Dim i As Integer
    On Local Error Resume Next
    For i = 1 To wbrPrincipal.UBound
        wbrPrincipal(i).Refresh
    Next
End Sub

Private Sub mnuUrlFavoritos_Click(index As Integer)
Dim Path As String
    
    'Obtiene la url
    Path = cFavoritos.ObtenerPath(index)
    If wbrPrincipal.Count <> 1 Then
        Wbr.Navigate Path
    Else
        AbrirNuevo Path
    End If
End Sub




Private Sub mnuVerCodigofuente_Click()
    On Error Resume Next
    If Not Wbr Is Nothing Then
        FrmCodigo_Fuente.Show , Me
        FrmCodigo_Fuente.Text1 = Wbr.Document.documentElement.OuterHTML
    End If
End Sub

Private Sub mnuVerCodigoFuenteP2_Click()
    mnuVerCodigofuente_Click
End Sub

Private Sub pbxContenedor_Resize()
Dim i As Long
    'Redimensionar los controles al área cliente del formulario
    If WindowState <> vbMinimized Then
        TS.Move 30, 30, pbxContenedor.ScaleWidth - 60, pbxContenedor.ScaleHeight - 60
        On Local Error Resume Next
        For i = wbrPrincipal.LBound To wbrPrincipal.UBound
            wbrPrincipal(i).Move 90, 390, pbxContenedor.ScaleWidth - 195, pbxContenedor.ScaleHeight - 480
        Next i
    End If
End Sub

Private Sub TS_BeforeClick(Cancel As Integer)
    Wbr.Visible = False 'Ocultar el «WebBrowser» activo

End Sub

Private Sub TS_Click()
    On Error Resume Next
    With Wbr
        cboURL.Text = .LocationURL
        .Visible = True 'Mostrar el «WebBrowser» de la pestaña seleccionada
        .ZOrder vbBringToFront
        Me.Caption = Wbr.LocationName
    Dim i As Integer
    For i = 1 To TS.Tabs.Count
        TS.Tabs(i).Image = 0
    Next
        TS.SelectedItem.Image = 14
    End With
        On Local Error Resume Next
        Wbr.SetFocus
        On Error GoTo 0
End Sub

Private Sub TS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
       Me.PopupMenu Me.mPopup
    End If
End Sub
    
Private Sub tbrPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    
    
    Select Case Button.Key
        Case "nue": mnuArchivoNuevo_Click
               Case "atr"
            On Error Resume Next
            Wbr.GoBack
                If Err Then
                    tbrPrincipal.Buttons("atr").Enabled = False
                    Err = 0
                End If
        Case "ade"
                On Error Resume Next
                Wbr.GoForward
                If Err Then
                    tbrPrincipal.Buttons("ade").Enabled = False
                    Err = 0
                End If
        Case "fav": mnuAgregarFavoritos_Click
        Case "det"
            On Error Resume Next
            Wbr.Stop
        Case "act"
            RefrescarPagina
        Case "ini"
            On Error Resume Next
            Wbr.Navigate Pagina
        Case "bus"
            On Error Resume Next
            Wbr.GoSearch
        Case "imp"
            On Error Resume Next
            mnuArchivoImprimir_Click
        End Select
        
        On Error GoTo 0
        
End Sub

Private Sub TS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Wbr Is Nothing Then
        TS.SelectedItem.ToolTipText = Wbr.LocationURL
    End If
End Sub

Private Sub wbrPrincipal_CommandStateChange(index As Integer, ByVal Command As Long, ByVal Enable As Boolean)
    'Activar o desactivar los botones Adelante/Atrás cuando cambia su estado de disponibilidad
        
        If TS.SelectedItem Is Nothing Then Exit Sub
        
        If TS.SelectedItem.Tag = index Then
            Select Case Command
                Case CSC_UPDATECOMMANDS
                   tbrPrincipal.Buttons("ade").Enabled = True
                   btnBack.Enabled = True
                Case CSC_NAVIGATEFORWARD
                    'tbrPrincipal.Buttons("ade").Enabled = Enable
                Case CSC_NAVIGATEBACK
                    'tbrPrincipal.Buttons("atr").Enabled = Enable
               ' Case -1
               '     tbrPrincipal.Buttons("atr").Enabled = False
               '     tbrPrincipal.Buttons("ade").Enabled = False
            End Select
        End If

End Sub

Private Sub wbrPrincipal_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    
    Dim frm_web As frmPrincipalBrowser
    'Nueva instancia del formulario
    Set frm_web = New frmPrincipalBrowser
    
    Set ppDisp = frm_web.Wbr.object
    
    frm_web.Show
    
End Sub

Private Sub wbrPrincipal_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo men

    pb.Max = ProgressMax
    pb.Value = Progress
    If ProgressMax <= 0 Then
       pb.Visible = False
    Else
       pb.Visible = True
    End If

Exit Sub

'Error
men:

If Err.Number = 380 Then Resume Next
End Sub

Private Sub wbrPrincipal_TitleChange(index As Integer, ByVal Text As String)
    'Actualizar el título de la pestaña con el título de la página actual
    'TS.Tabs(index).Caption = Text
    TS.SelectedItem.Caption = " " & Left(Text, 25) & " ..."
    Me.Caption = Wbr.LocationName + " - Mammoth"
End Sub

Private Sub wbrPrincipal_NavigateComplete2(index As Integer, ByVal pDisp As Object, url As Variant)
    'Actualizar la URL mostrada en la barra de direcciones con la URL de la página actual
    If Not TS.SelectedItem Is Nothing Then
        If TS.SelectedItem.Tag = index Then
            cboURL.Text = wbrPrincipal(index).LocationURL
              cboURL.BackColor = vbWhite
              cboURL.Font.Bold = False
            

        End If
    End If
End Sub

Private Sub wbrPrincipal_StatusTextChange(index As Integer, ByVal Text As String)
    'Actualizar la barra de estado con el de la página actual
    If Not TS.SelectedItem Is Nothing Then
        If TS.SelectedItem.Tag = index Then stbPrincipal.SimpleText = Text
    End If
End Sub

' para los accesos directos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub AgregarAccesoDirecto(url As String)
    
    
    'Nuevo acceso directo en el registro y en el menú
    Call c_Recientes.NuevoElemento(url)
    
End Sub

' Al hacer click en el acceso directo abre el archivo _
  en un nuevo formulario cargando el contenido en el RichTextBox

Private Sub mnuReciente_Click(index As Integer)
Dim Path As String
    
    'Obtiene el path del archivo
    Path = c_Recientes.ObtenerPath(index)
    If wbrPrincipal.Count <> 1 Then
        Wbr.Navigate Path
    Else
        AbrirNuevo Path
    End If
End Sub
 _
 _
 _
 _
 _
 _
 _

