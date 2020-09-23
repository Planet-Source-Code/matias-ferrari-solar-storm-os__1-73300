VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDesktop 
   BorderStyle     =   0  'None
   ClientHeight    =   12015
   ClientLeft      =   180
   ClientTop       =   705
   ClientWidth     =   21570
   Icon            =   "FrmDesktop.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12015
   ScaleWidth      =   21570
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox FrmUnidades 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   6840
      ScaleHeight     =   3585
      ScaleWidth      =   2025
      TabIndex        =   37
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         MouseIcon       =   "FrmDesktop.frx":0CCA
         TabIndex        =   39
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":0E1C
         Picture         =   "FrmDesktop.frx":0F6E
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Accesos"
      Top             =   10680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame FrmAccesorios 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   9480
      TabIndex        =   27
      Top             =   4560
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label lblWorpad 
         BackStyle       =   0  'Transparent
         Caption         =   "Worpad del Sistema Operativo"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":1C38
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Image ImgWorpad 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":1D8A
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":1EDC
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image ImgVolumenes 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":27A6
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":28F8
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label LblVolumenes 
         BackStyle       =   0  'Transparent
         Caption         =   "Control de Volumenes"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":35C2
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         Height          =   4455
         Left            =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label LblBlockdeNotas 
         BackStyle       =   0  'Transparent
         Caption         =   "Block de Notas"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":3714
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   240
         Width           =   2295
      End
      Begin VB.Image ImgBloackdeNotas 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":3866
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":39B8
         Top             =   120
         Width           =   480
      End
      Begin VB.Label LblCalculadora 
         BackStyle       =   0  'Transparent
         Caption         =   "Calculadora"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":4282
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   720
         Width           =   2295
      End
      Begin VB.Image ImgCalculadora 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":43D4
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":4526
         Top             =   600
         Width           =   480
      End
      Begin VB.Label LblRemoteDesktop 
         BackStyle       =   0  'Transparent
         Caption         =   "Escritorio Remoto"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":4DF0
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Image ImgRemoteDesktop 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":4F42
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":5094
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image ImgExplorer 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":595E
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":5AB0
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label LblExplorer 
         BackStyle       =   0  'Transparent
         Caption         =   "Explorador de Solar Storm"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":637A
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Image ImgLibreta 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":64CC
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":661E
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label LblLibretadeDirecciones 
         BackStyle       =   0  'Transparent
         Caption         =   "Libreta de Direcciones"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":6EE8
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Image ImgPaint 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":703A
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":718C
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label LblPaint 
         BackStyle       =   0  'Transparent
         Caption         =   "Paint del Sistema Operativo"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":7A56
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label LblDesfrag 
         BackStyle       =   0  'Transparent
         Caption         =   "Desfragmentador de Discos"
         Height          =   255
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":7BA8
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Image ImgDesfrag 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":7CFA
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":7E4C
         Top             =   3000
         Width           =   480
      End
   End
   Begin VB.PictureBox PicPostIT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   18720
      ScaleHeight     =   2625
      ScaleWidth      =   2625
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Data Data1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Post-It"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         DataField       =   "Post"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   2655
      End
      Begin VB.Image ImgEliminarPost 
         Height          =   240
         Left            =   1920
         MouseIcon       =   "FrmDesktop.frx":8716
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":8868
         ToolTipText     =   "Eliminar Post"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image ImgCerrarPosts 
         Height          =   240
         Left            =   2400
         MouseIcon       =   "FrmDesktop.frx":8BF2
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":8D44
         ToolTipText     =   "Cerrar Post-It"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image ImgAddPost 
         Height          =   240
         Left            =   1560
         MouseIcon       =   "FrmDesktop.frx":90CE
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":9220
         ToolTipText     =   "Agregar Post"
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblFecha_PostIT 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         DataField       =   "Fecha"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   6100
      Left            =   18600
      Top             =   0
   End
   Begin VB.Frame FrmPrograms 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   7440
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Image ImgVisorDeImagenes 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":95AA
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":96FC
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblVisorDeImagenes 
         BackStyle       =   0  'Transparent
         Caption         =   "Visor de Imagenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":A3C6
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Image ImgMostrarProcesos 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":A518
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":A66A
         Top             =   4320
         Width           =   480
      End
      Begin VB.Label LblMostrarProcesos 
         BackStyle       =   0  'Transparent
         Caption         =   "Mostrar Procesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":B334
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label LblCMD 
         BackStyle       =   0  'Transparent
         Caption         =   "Simbolo de Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":B486
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Image ImgCMD 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":B5D8
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":B72A
         Top             =   3720
         Width           =   480
      End
      Begin VB.Label lblPostIT 
         BackStyle       =   0  'Transparent
         Caption         =   "Post - It"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":C3F4
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Image ImgPostIT 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":C546
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":C698
         Top             =   2520
         Width           =   480
      End
      Begin VB.Image ImgHerramientas 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":CF62
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":D0B4
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label lblHerramientas 
         BackStyle       =   0  'Transparent
         Caption         =   "Herramientas                >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":DD7E
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         Height          =   4815
         Left            =   0
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label LblPortables 
         BackStyle       =   0  'Transparent
         Caption         =   "Portables                     >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":DED0
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Image ImgPortables 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":E022
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":E174
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label lblAccesories 
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios                   >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":EE3E
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image ImgAccesories 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":EF90
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":F0E2
         Top             =   120
         Width           =   480
      End
      Begin VB.Label LblGames 
         BackStyle       =   0  'Transparent
         Caption         =   "Juegos                        >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":FDAC
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   840
         Width           =   2295
      End
      Begin VB.Image ImgGames 
         Height          =   480
         Left            =   120
         MouseIcon       =   "FrmDesktop.frx":FEFE
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":10050
         Top             =   720
         Width           =   480
      End
   End
   Begin MSComCtl2.MonthView Calendario 
      Height          =   2370
      Left            =   18960
      TabIndex        =   13
      Top             =   9240
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   16711682
      CurrentDate     =   40144
   End
   Begin VB.Timer Tmrwallpaper 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   18240
      Top             =   0
   End
   Begin VB.Timer TmrLoad 
      Interval        =   50
      Left            =   17880
      Top             =   0
   End
   Begin VB.Timer TmrDesktop 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   17520
      Top             =   0
   End
   Begin VB.Frame FrmStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label LblBuscar 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar                                            >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   41
         Top             =   4560
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         Height          =   6735
         Left            =   0
         Top             =   0
         Width           =   4815
      End
      Begin VB.Image ImgAbout 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":10D1A
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":10E6C
         Top             =   1320
         Width           =   480
      End
      Begin VB.Label LblAbout 
         BackStyle       =   0  'Transparent
         Caption         =   "Acerca De...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":11B36
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Image ImgUnidades 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":11C88
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":11DDA
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label LblUnidades 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades                                        >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":12AA4
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   840
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Image ImgPrograms 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":12BF6
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":12D48
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label LblPrograms 
         BackStyle       =   0  'Transparent
         Caption         =   "Programas                                       >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":13A12
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":13B64
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":13CB6
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Descargar Actualizaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":14980
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   840
         Width           =   3255
      End
      Begin VB.Image ImgInternet 
         Height          =   720
         Left            =   720
         MouseIcon       =   "FrmDesktop.frx":14AD2
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":14C24
         Top             =   0
         Width           =   720
      End
      Begin VB.Label LblInternet 
         BackStyle       =   0  'Transparent
         Caption         =   "Mamooth Browser"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":168EE
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label LblConfiguracion 
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración                                  >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":16A40
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Image imgConfiguracion 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":16B92
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":16CE4
         Top             =   3720
         Width           =   480
      End
      Begin VB.Image ImgBuscar 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":179AE
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":17B00
         Top             =   4440
         Width           =   480
      End
      Begin VB.Label LblRun 
         BackStyle       =   0  'Transparent
         Caption         =   "Ejecutar...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":187CA
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Image ImgRun 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":1891C
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":18A6E
         Top             =   5160
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   720
         X2              =   4440
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label lblOpciones 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Que Desea Hacer?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   7110
         Left            =   0
         Picture         =   "FrmDesktop.frx":19738
         Stretch         =   -1  'True
         Top             =   0
         Width           =   555
      End
      Begin VB.Image BtnCerrarSesion 
         Height          =   480
         Left            =   840
         MouseIcon       =   "FrmDesktop.frx":19E98
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":19FEA
         Top             =   6120
         Width           =   480
      End
      Begin VB.Image BtnRestart 
         Height          =   480
         Left            =   1440
         MouseIcon       =   "FrmDesktop.frx":2B4EC
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":2B63E
         Top             =   6120
         Width           =   480
      End
      Begin VB.Image BtnShutDown 
         Height          =   480
         Left            =   2040
         MouseIcon       =   "FrmDesktop.frx":2C308
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":2C45A
         Top             =   6120
         Width           =   480
      End
   End
   Begin VB.PictureBox PicToolbar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   21570
      TabIndex        =   0
      Top             =   11430
      Width           =   21570
      Begin VB.Image ImgRed 
         Height          =   480
         Left            =   16680
         MouseIcon       =   "FrmDesktop.frx":2D124
         MousePointer    =   99  'Custom
         Top             =   0
         Width           =   480
      End
      Begin VB.Image BtnStart 
         Height          =   720
         Left            =   0
         MouseIcon       =   "FrmDesktop.frx":2D276
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":2D3C8
         Top             =   -50
         Width           =   720
      End
      Begin VB.Image BtnMultimedia 
         Height          =   480
         Left            =   17280
         MouseIcon       =   "FrmDesktop.frx":2F092
         MousePointer    =   99  'Custom
         Picture         =   "FrmDesktop.frx":2F1E4
         Top             =   60
         Width           =   480
      End
      Begin VB.Label LblHora 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 AM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   17880
         MouseIcon       =   "FrmDesktop.frx":2FEAE
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   180
         Width           =   1695
      End
      Begin VB.Image ImgToolBar 
         Height          =   690
         Left            =   0
         Picture         =   "FrmDesktop.frx":30000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   21720
      End
   End
   Begin VB.Label LblBrowser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mamooth Browser"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      MouseIcon       =   "FrmDesktop.frx":308C1
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Image ImgBrowser 
      Height          =   720
      Left            =   240
      MouseIcon       =   "FrmDesktop.frx":30A13
      MousePointer    =   99  'Custom
      Picture         =   "FrmDesktop.frx":30B65
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label LblMyPortables 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Portables"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      MouseIcon       =   "FrmDesktop.frx":3282F
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image ImgMyPortables 
      Height          =   720
      Left            =   240
      MouseIcon       =   "FrmDesktop.frx":32981
      MousePointer    =   99  'Custom
      Picture         =   "FrmDesktop.frx":32AD3
      Top             =   2400
      Width           =   720
   End
   Begin VB.Image ImgRedOff 
      Height          =   480
      Left            =   20160
      Picture         =   "FrmDesktop.frx":3479D
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgRedON 
      Height          =   480
      Left            =   19560
      Picture         =   "FrmDesktop.frx":35467
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgMyComputer 
      Height          =   720
      Left            =   240
      MouseIcon       =   "FrmDesktop.frx":36131
      MousePointer    =   99  'Custom
      Picture         =   "FrmDesktop.frx":36283
      Top             =   0
      Width           =   720
   End
   Begin VB.Label LblMyComputer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Computer"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      MouseIcon       =   "FrmDesktop.frx":37F4D
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image BtnDouments 
      Height          =   720
      Left            =   240
      MouseIcon       =   "FrmDesktop.frx":3809F
      MousePointer    =   99  'Custom
      Picture         =   "FrmDesktop.frx":381F1
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label LblDocuments 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My Documents"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      MouseIcon       =   "FrmDesktop.frx":39EBB
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image imgWall 
      Height          =   10320
      Left            =   1680
      Stretch         =   -1  'True
      Top             =   120
      Width           =   15240
   End
   Begin VB.Image BtnStartClick 
      Height          =   720
      Left            =   6720
      Picture         =   "FrmDesktop.frx":3A00D
      Top             =   10680
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image BtnStartNormal 
      Height          =   720
      Left            =   5880
      Picture         =   "FrmDesktop.frx":3BCD7
      Top             =   10680
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "FrmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long   'Contador.
Dim Tiempo As String  'Tiempo total transcurrido.

'Constantes para determinar que tipo de Red estamos conectados

Const NETWORK_ALIVE_AOL = &H4
Const NETWORK_ALIVE_LAN = &H1
Const NETWORK_ALIVE_WAN = &H2

'Función Api IsNetworkAlive para detectar _
 si estamos conectados y a que tipo de red
Private Declare Function IsNetworkAlive Lib "SENSAPI.DLL" ( _
    ByRef lpdwFlags As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long





Private Sub BtnCerrarSesion_Click()
TmrDesktop.Enabled = True
FrmLoghon.Show

End Sub

Private Sub BtnCerrarSesion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpciones.Caption = "Cerrar Sesión"
End Sub

Private Sub BtnDouments_Click()
Me.TmrDesktop.Enabled = True

End Sub

Private Sub BtnDouments_DblClick()
Shell "explorer.exe " & App.Path & "\Documents", vbNormalFocus
Me.TmrDesktop.Enabled = True

End Sub

Private Sub BtnMultimedia_Click()
Shell "sndvol32.exe", vbNormalFocus
Me.TmrDesktop.Enabled = True

End Sub

Private Sub BtnRestart_Click()
FrmSplash.Show
FrmSplash.Timer1.Enabled = True
Unload Me

End Sub

Private Sub BtnRestart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpciones.Caption = "Reiniciar Software"
End Sub

Private Sub BtnShutDown_Click()
  Dim StartWindow As Long ' Lo primero que tenemos que hacer es localizar la barra de tareas con la instrucción ' de debajo y luego con el manejador pasarsela a la función que la oculta o la muestra
    StartWindow = FindWindow("Shell_TrayWnd", vbNullString)

ShowWindow StartWindow, 1& ' La mostramos de nuevo la barra de tareas

  End
End Sub

Private Sub BtnShutDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOpciones.Caption = "Apagar USB"
End Sub

Private Sub BtnStart_Click()
BtnStart.Picture = BtnStartClick.Picture
FrmStart.Visible = True
End Sub



Private Sub Form_Load()

i = 0 'Inicializar el contador.
Timer1.Interval = 0    'Detener el cronometro
Timer1.Interval = 1    'Iniciar el cronometro


' ESTA FUNCION OCULTA LA BARRA DE TAREAS DE WINDOWS
  Dim StartWindow As Long ' Lo primero que tenemos que hacer es localizar la barra de tareas con la instrucción ' de debajo y luego con el manejador pasarsela a la función que la oculta o la muestra
    StartWindow = FindWindow("Shell_TrayWnd", vbNullString)

        ShowWindow StartWindow, 0&  ' La ocultamos




LblHora.Caption = Format(Now, "hh:mm")
Tmrwallpaper.Enabled = True
Me.ImgToolBar.Width = Me.PicToolbar.Width



 Dim ret As Long

    'Si la Api retorna 0 quiere decir que no hay ningun tipo de conexión de Red
        If IsNetworkAlive(ret) = 0 Then

            ImgRed.Picture = ImgRedOff.Picture
            ImgRed.ToolTipText = "Conexion a Internet OFF"
        Else
            ' hay conexión , y muestra el tipo
            ImgRed.Picture = ImgRedON.Picture
            ImgRed.ToolTipText = "Conexion a Internet ON"
  
    End If

Data1.DatabaseName = App.Path & "\BD.ss"
Data2.DatabaseName = App.Path & "\BD.ss"

End Sub








Private Sub ImgAbout_Click()
LblAbout_Click
End Sub

Private Sub ImgAccesories_Click()
lblAccesories_Click
End Sub

Private Sub ImgAddPost_Click()
On Error Resume Next
Data1.Recordset.AddNew
lblFecha_PostIT.Caption = Date
End Sub

Private Sub ImgBloackdeNotas_Click()
LblBlockdeNotas_Click
End Sub

Private Sub ImgBrowser_DblClick()
FrmSplashBrowser.Show
Me.TmrDesktop.Enabled = True

End Sub



Private Sub ImgBuscar_Click()
LblBuscar_Click
End Sub

Private Sub ImgCalculadora_Click()
LblCalculadora_Click
End Sub

Private Sub ImgCerrarPosts_Click()
PicPostIT.Visible = False
End Sub

Private Sub ImgCMD_Click()
LblCMD_Click
End Sub

Private Sub imgConfiguracion_Click()
LblConfiguracion_Click
End Sub

Private Sub ImgDesfrag_Click()
LblDesfrag_Click
End Sub

Private Sub ImgEliminarPost_Click()
On Error Resume Next
Data1.Recordset.Delete
Data1.Refresh
End Sub

Private Sub ImgExplorer_Click()
LblExplorer_Click
End Sub

Private Sub ImgGames_Click()
lblGames_Click
End Sub

Private Sub ImgHerramientas_Click()
lblHerramientas_Click
End Sub

Private Sub ImgLibreta_Click()
LblLibretadeDirecciones_Click
End Sub

Private Sub ImgMostrarProcesos_Click()
LblMostrarProcesos_Click
End Sub



Private Sub ImgMyComputer_Click()
Me.TmrDesktop.Enabled = True

End Sub

Private Sub ImgMyComputer_DblClick()
Shell "explorer.exe /n,  /select, C:\", vbNormalFocus
Me.TmrDesktop.Enabled = True


End Sub

Private Sub ImgMyPortables_Click()
Me.TmrDesktop.Enabled = True

End Sub

Private Sub ImgMyPortables_DblClick()
LblMyPortables_DblClick
End Sub

Private Sub ImgPaint_Click()
LblPaint_Click
End Sub

Private Sub ImgPortables_Click()
LblPortables_Click
End Sub

Private Sub ImgPostIT_Click()
lblPostIT_Click
End Sub

Private Sub ImgPrograms_Click()
LblPrograms_Click
End Sub

Private Sub ImgRed_Click()
FrmRed.Show
FrmRed.Left = Me.BtnMultimedia.Left - 3400
FrmRed.Top = PicToolbar.Top - 3550

End Sub

Private Sub ImgRemoteDesktop_Click()
LblRemoteDesktop_Click
End Sub



Private Sub ImgRun_Click()
FrmRun.Show
End Sub

Private Sub ImgToolBar_Click()
TmrDesktop.Enabled = True
End Sub


Private Sub ImgVisorDeImagenes_Click()
FrmImageViewer.Show
TmrDesktop.Enabled = True

End Sub

Private Sub ImgVolumenes_Click()
LblVolumenes_Click
End Sub

Private Sub imgWall_Click()
TmrDesktop.Enabled = True

End Sub











Private Sub ImgWorpad_Click()
lblWorpad_Click
End Sub







Private Sub LblAbout_Click()
FrmAbout.Show
Me.TmrDesktop.Enabled = True

End Sub

Private Sub lblAccesories_Click()
Me.FrmAccesorios.Left = Me.FrmStart.Left + 7830
Me.FrmAccesorios.Top = PicToolbar.Top - 4800
Me.FrmAccesorios.Visible = True

End Sub



Private Sub LblBlockdeNotas_Click()
MDINotepad.Show

Me.TmrDesktop.Enabled = True

End Sub

Private Sub LblBrowser_DblClick()
ImgBrowser_DblClick
End Sub


Private Sub LblBuscar_Click()
FrmBuscar.Show
TmrDesktop.Enabled = True

End Sub

Private Sub LblCalculadora_Click()
FrmCalculadora.Show
TmrDesktop.Enabled = True

End Sub

Private Sub LblCMD_Click()
On Error Resume Next
Shell "cmd.exe ", vbNormalFocus
TmrDesktop.Enabled = True

End Sub

Private Sub LblConfiguracion_Click()
FrmConfiguracion.Show
Me.TmrDesktop.Enabled = True
End Sub

Private Sub LblDesfrag_Click()
On Error Resume Next
        Buffer = Shell(Environ("WinDir") & "\system32\mmc.exe " & Environ("WinDir") & "\system32\dfrg.msc c:", vbNormalFocus) ' Ejecuta el Desfragmentador
        Me.TmrDesktop.Enabled = True

End Sub

Private Sub LblDocuments_DblClick()
BtnDouments_DblClick
End Sub

Private Sub LblExplorer_Click()
Shell App.Path & "\explorerss.exe ", vbNormalFocus
Me.TmrDesktop.Enabled = True
End Sub

Private Sub lblGames_Click()
FrmGames.Show
Me.TmrDesktop.Enabled = True
End Sub

Private Sub lblHerramientas_Click()
Me.TmrDesktop.Enabled = True
FrmTools.Show
End Sub

Private Sub LblHora_Click()
Calendario.Visible = False
End Sub

Private Sub LblHora_DblClick()
Calendario.Left = Me.BtnMultimedia.Left - 1000
Calendario.Top = PicToolbar.Top - 2380
Calendario.Visible = True
Calendario.Value = Date

End Sub


Private Sub LblInternet_Click()
FrmSplashBrowser.Show
Me.TmrDesktop.Enabled = True
End Sub

Private Sub LblLibretadeDirecciones_Click()

FrmLibretadeDirecciones.Show
Me.TmrDesktop.Enabled = True

End Sub

Private Sub LblMostrarProcesos_Click()
FrmProcesos.Show
TmrDesktop.Enabled = True

End Sub





Private Sub LblMyComputer_DblClick()
ImgMyComputer_DblClick
End Sub

Private Sub LblMyPortables_DblClick()
FrmPortables.Show
Me.TmrDesktop.Enabled = True

End Sub

Private Sub LblPaint_Click()
Shell "mspaint.exe ", vbNormalFocus
Me.TmrDesktop.Enabled = True

End Sub

Private Sub LblPortables_Click()
FrmPortables.Show
Me.TmrDesktop.Enabled = True

End Sub

Private Sub lblPostIT_Click()
Me.TmrDesktop.Enabled = True
PicPostIT.Left = FrmDesktop.Width - 2850
PicPostIT.Top = FrmDesktop.Top + 550


PicPostIT.Visible = True
Data1.DatabaseName = App.Path & "\BD.ss"

End Sub

Private Sub LblPrograms_Click()
Me.FrmPrograms.Left = Me.FrmStart.Left + 4815
Me.FrmPrograms.Top = PicToolbar.Top - 4800
Me.FrmPrograms.Visible = True
Me.FrmUnidades.Visible = False
End Sub

Private Sub LblRemoteDesktop_Click()
Shell "mstsc.exe ", vbNormalFocus
Me.TmrDesktop.Enabled = True

End Sub


Private Sub LblRun_Click()
FrmRun.Show
End Sub



Private Sub LblUnidades_Click()
Me.FrmUnidades.Left = Me.FrmStart.Left + 4815
Me.FrmUnidades.Top = PicToolbar.Top - 3600
Me.FrmUnidades.Visible = True
FrmPrograms.Visible = False
Me.FrmAccesorios.Visible = False


List1.Clear

GetDrives List1, REMOVABLE
GetDrives List1, CDROM
GetDrives List1, Fixed
GetDrives List1, RAMDISK
GetDrives List1, REMOTE
End Sub

Private Sub lblVisorDeImagenes_Click()
ImgVisorDeImagenes_Click
End Sub

Private Sub LblVolumenes_Click()
Shell "sndvol32.exe", vbNormalFocus
Me.TmrDesktop.Enabled = True

End Sub

Private Sub lblWorpad_Click()
On Error Resume Next
Shell "wordpad.exe", vbNormalFocus
Me.TmrDesktop.Enabled = True


End Sub



Private Sub List1_DblClick()
On Error Resume Next
FrmPropiedadesdeUnidad.Show
FrmPropiedadesdeUnidad.Drive1.Drive = List1.Text
FrmPropiedadesdeUnidad.Text1.Text = List1.Text


End Sub

Private Sub Timer1_Timer()
i = i + 1
Tiempo = Format(Int(i / 6000) Mod 24, "00") & ":" & _
         Format(Int(i / 600) Mod 60, "00") & ":" & _
         Format(Int(i / 10) Mod 60, "00") & ":" & _
         Format(i Mod 10, "00")
FrmRed.LblDuración.Caption = Tiempo

End Sub

Private Sub TmrDesktop_Timer()
BtnStart.Picture = BtnStartNormal.Picture
FrmStart.Visible = False
FrmPrograms.Visible = False
Me.FrmAccesorios.Visible = False
Me.FrmUnidades.Visible = False


TmrDesktop.Enabled = False
lblOpciones.Caption = "Que Desea Hacer?"
Calendario.Visible = False
End Sub

Private Sub TmrLoad_Timer()
Me.LblHora.Left = Me.PicToolbar.Width - 1000
Me.BtnMultimedia.Left = Me.PicToolbar.Width - 1590
Me.ImgRed.Left = Me.PicToolbar.Width - 2180

Me.FrmStart.Top = PicToolbar.Top - 6740
Me.FrmStart.Left = 10

LblHora.Caption = Format(Now, "hh:mm")
Me.ImgToolBar.Width = Me.PicToolbar.Width
 Dim ret As Long

    'Si la Api retorna 0 quiere decir que no hay ningun tipo de conexión de Red
        If IsNetworkAlive(ret) = 0 Then

            ImgRed.Picture = ImgRedOff.Picture
            ImgRed.ToolTipText = "Conexion a Internet OFF"
        Else
            ' hay conexión , y muestra el tipo
            ImgRed.Picture = ImgRedON.Picture
            ImgRed.ToolTipText = "Conexion a Internet ON"
  
    End If

TmrLoad.Interval = 12480

End Sub

Private Sub Tmrwallpaper_Timer()
Me.imgWall.Top = 0
Me.imgWall.Left = 0
Me.imgWall.Width = FrmDesktop.Width
Me.imgWall.Height = FrmDesktop.Height - Me.ImgToolBar.Height + 200
Tmrwallpaper.Enabled = False

End Sub
 _
 _
 _
 _
 _
 _
 _
 _
 _
 _

