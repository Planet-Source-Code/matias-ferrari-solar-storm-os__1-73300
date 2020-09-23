VERSION 5.00
Begin VB.Form FrmBuscar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Archivos"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   2985
      TabIndex        =   19
      Top             =   720
      Width           =   3015
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Archivos "
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
         Left            =   600
         TabIndex        =   20
         Top             =   120
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "FrmBuscar.frx":0CCA
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4665
      ScaleWidth      =   2985
      TabIndex        =   8
      Top             =   1320
      Width           =   3015
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   1440
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   3720
         Width           =   1215
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Archivos en :"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre o Extension a Buscar :"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cambiar Unidad :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   2760
         Y1              =   3120
         Y2              =   3120
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10395
      TabIndex        =   1
      Top             =   0
      Width           =   10425
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "Buscar...."
         Top             =   120
         Width           =   8055
      End
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   3240
         ScaleHeight     =   15
         ScaleWidth      =   5295
         TabIndex        =   2
         Top             =   480
         Width           =   5295
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   9840
         Picture         =   "FrmBuscar.frx":1994
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5100
      Left            =   3240
      TabIndex        =   0
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Archivos "
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
      TabIndex        =   18
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "FrmBuscar.frx":265E
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6120
      Width           =   3615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   6120
      Width           =   7095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Archivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ruta :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "FrmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'***************************************************************************
'*  Controles         : Command1 ( para buscar) _
                        Text1 ( para indicar el Path) _
                        Text2 ( para los archivos, por ejemplo *.txt ) _
                        List1
'***************************************************************************

Private Sub Command1_Click()

    Dim Path As String
    Dim Pattern As String
    Dim FileSize As Currency
    Dim Count_Archivos As Long
    Dim Count_Dir As Long

    Screen.MousePointer = vbHourglass
    
    'Borramos el contenido del List1
    List1.Clear
    
    'Path y archivos a buscar
    Path = Text1.Text
    Pattern = Text2.Text
    
    'Llamamos a la función para buscar y que nos retorne algunos datos
    FileSize = FindFilesAPI(Path, Pattern, _
                            Count_Archivos, _
                            Count_Dir, List1)

    'Mostramos los resultados
    
    'Cantidad de archivos encontrados
    Label7.Caption = Count_Archivos & " Archivos encontrados en " & Count_Dir + 1 & " Directorios"
    
    'Tamaño Total en Bytes de los archivos encontrados
    Label8.Caption = "Tamaño total de los archivos: " & _
            Path & " = " & _
            Format(FileSize, "#,###,###,##0") & " Bytes"

    Screen.MousePointer = vbDefault

End Sub


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Drive1_Change()
Text1.Text = Drive1.Drive
End Sub

Private Sub Form_Load()
    'Directorio de windows
    Text1.Text = Environ("WinDir")
    'Archivos txt
    Text2.Text = "*.txt"
    
    Command1.Caption = "  >> Buscar "

End Sub


Private Sub List1_DblClick()
Me.txtPath.Text = List1.Text
End Sub

Private Sub txtPath_Change()
On Error Resume Next
Shell (txtPath.Text), vbNormalFocus

End Sub
