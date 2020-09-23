VERSION 5.00
Begin VB.Form FrmWallpapers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar Wallpaper"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5040
      Top             =   5880
   End
   Begin VB.TextBox TxtWallActual 
      DataField       =   "Wallpaper"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Wallpaper"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5280
      Width           =   5535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aplicar"
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8400
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox TxtWallpaper 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   3975
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   5880
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Wallpaper :"
      Height          =   255
      Left            =   5880
      TabIndex        =   8
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ruta del Wallpaper :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Wallpapers Disponibles :"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
   Begin VB.Image ImgPreview 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   4455
      Left            =   120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vista Previa :"
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "FrmWallpapers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Me.TxtWallActual.Text = Me.TxtWallpaper.Text
FrmDesktop.imgWall.Picture = LoadPicture(Text1.Text)
Unload Me
End Sub

Private Sub File1_Click()

TxtWallpaper.Text = File1.FileName
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\BD.ss"
File1.Path = App.Path & "\Wallpapers"
Me.Timer1.Enabled = True
End Sub

Private Sub Text1_Change()
On Error Resume Next

Me.ImgPreview.Picture = LoadPicture(Text1.Text)

End Sub

Private Sub Timer1_Timer()
Me.Timer1.Enabled = False
Me.TxtWallpaper.Text = Me.TxtWallActual.Text

End Sub

Private Sub TxtWallpaper_Change()
Text1.Text = App.Path & "\Wallpapers\" & TxtWallpaper.Text

End Sub
