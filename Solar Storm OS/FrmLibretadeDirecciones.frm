VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmLibretadeDirecciones 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libreta de Direcciones"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   Icon            =   "FrmLibretadeDirecciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      DataField       =   "Notas"
      DataSource      =   "Data1"
      Height          =   1485
      Left            =   5160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3720
      Width           =   5895
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "Fecha nacimiento"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   15
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      DataField       =   "DirCorreoElectrónico"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   13
      Top             =   2640
      Width           =   5895
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      DataField       =   "TeléfonoMóvil"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8760
      TabIndex        =   11
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      DataField       =   "TeléfonoCasa"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "Ciudad"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Dirección"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "Apellidos"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8760
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Nombre"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Lista de direcciones"
      Top             =   5640
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmLibretadeDirecciones.frx":08CA
      Height          =   5775
      Left            =   120
      OleObjectBlob   =   "FrmLibretadeDirecciones.frx":08DE
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label LblEliminarUsuario 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7560
      MouseIcon       =   "FrmLibretadeDirecciones.frx":12C9
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5640
      Width           =   735
   End
   Begin VB.Image ImgEliminarUsuario 
      Height          =   480
      Left            =   7080
      MouseIcon       =   "FrmLibretadeDirecciones.frx":141B
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibretadeDirecciones.frx":156D
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label LblCrearUsuario 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8880
      MouseIcon       =   "FrmLibretadeDirecciones.frx":1E37
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   5640
      Width           =   735
   End
   Begin VB.Image IgCrearUsuario 
      Height          =   480
      Left            =   8520
      MouseIcon       =   "FrmLibretadeDirecciones.frx":1F89
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibretadeDirecciones.frx":20DB
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image ImgCancelarUsuario 
      Height          =   480
      Left            =   9720
      MouseIcon       =   "FrmLibretadeDirecciones.frx":29A5
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibretadeDirecciones.frx":2AF7
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label lblCancelarUsuario 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abandonar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10080
      MouseIcon       =   "FrmLibretadeDirecciones.frx":33C1
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Notas :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cumpleaños :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Celular :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. Casa :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "FrmLibretadeDirecciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\BD.ss"

End Sub

Private Sub IgCrearUsuario_Click()
Data1.Recordset.AddNew
End Sub

Private Sub ImgCancelarUsuario_Click()
Unload Me
End Sub

Private Sub ImgEliminarUsuario_Click()
On Error Resume Next
Data1.Recordset.Delete
End Sub

Private Sub lblCancelarUsuario_Click()
Unload Me
End Sub

Private Sub LblCrearUsuario_Click()
IgCrearUsuario_Click
End Sub

Private Sub LblEliminarUsuario_Click()
ImgEliminarUsuario_Click
End Sub
