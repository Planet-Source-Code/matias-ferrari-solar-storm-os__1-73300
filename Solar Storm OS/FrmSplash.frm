VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   6250
      Left            =   5760
      Top             =   4800
   End
   Begin VB.TextBox TxtWallActual 
      DataField       =   "Wallpaper"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   4800
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
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Wallpaper"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1508
      _cy             =   1085
   End
   Begin VB.Image Image1 
      Height          =   5280
      Left            =   0
      Picture         =   "FrmSplash.frx":0000
      Top             =   0
      Width           =   8115
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\BD.ss"
WMP.url = App.Path & "/Sound/Start.wav"
WMP.Controls.Play
FrmDesktop.Show
FrmDesktop.Visible = False

End Sub


Private Sub Timer1_Timer()
Unload FrmSplash
Timer1.Enabled = False
WMP.Controls.Stop
FrmDesktop.imgWall.Picture = LoadPicture(App.Path & "\Wallpapers\" & TxtWallActual.Text)
FrmDesktop.Visible = True


End Sub
