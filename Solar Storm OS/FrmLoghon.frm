VERSION 5.00
Begin VB.Form FrmLoghon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Iniciar Sesion"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contraseña"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtCont 
      DataField       =   "Contraseña"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtContraseña 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Tmrwallpaper 
      Interval        =   25
      Left            =   10800
      Top             =   240
   End
   Begin VB.Image ImgAbout 
      Height          =   480
      Left            =   240
      MouseIcon       =   "FrmLoghon.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoghon.frx":0152
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label LblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acerca De..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      MouseIcon       =   "FrmLoghon.frx":0E1C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image ImgNext 
      Height          =   480
      Left            =   10680
      MouseIcon       =   "FrmLoghon.frx":0F6E
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoghon.frx":10C0
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IcoLoghon 
      Height          =   480
      Left            =   240
      MouseIcon       =   "FrmLoghon.frx":1D8A
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoghon.frx":1EDC
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblLoghon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Iniciar Sesion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      MouseIcon       =   "FrmLoghon.frx":2BA6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Image ImgLoghon 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   8280
      Picture         =   "FrmLoghon.frx":2CF8
      Top             =   2640
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image BtnRestart 
      Height          =   480
      Left            =   240
      MouseIcon       =   "FrmLoghon.frx":F542
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoghon.frx":F694
      Top             =   720
      Width           =   480
   End
   Begin VB.Label lblOpciones 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reiniciar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      MouseIcon       =   "FrmLoghon.frx":1035E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Image BtnShutDown 
      Height          =   480
      Left            =   240
      MouseIcon       =   "FrmLoghon.frx":104B0
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoghon.frx":10602
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LblShutDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apagar el Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      MouseIcon       =   "FrmLoghon.frx":112CC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image imgWall 
      Height          =   3120
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "FrmLoghon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub BtnRestart_Click()
FrmSplash.Show
FrmSplash.Timer1.Enabled = True
Unload Me
Unload FrmDesktop

End Sub

Private Sub BtnShutDown_Click()
  Dim StartWindow As Long ' Lo primero que tenemos que hacer es localizar la barra de tareas con la instrucción ' de debajo y luego con el manejador pasarsela a la función que la oculta o la muestra
    StartWindow = FindWindow("Shell_TrayWnd", vbNullString)

ShowWindow StartWindow, 1& ' La mostramos de nuevo la barra de tareas

  End

End Sub

Private Sub Form_Load()
On Error Resume Next
' ESTA FUNCION OCULTA LA BARRA DE TAREAS DE WINDOWS
  Dim StartWindow As Long ' Lo primero que tenemos que hacer es localizar la barra de tareas con la instrucción ' de debajo y luego con el manejador pasarsela a la función que la oculta o la muestra
    StartWindow = FindWindow("Shell_TrayWnd", vbNullString)

 imgWall.Picture = LoadPicture(App.Path & "\Wallpapers\Loghon.jpg")

Data1.DatabaseName = App.Path & "\BD.ss"
End Sub

Private Sub Image2_Click()

End Sub

Private Sub IcoLoghon_Click()
lblLoghon_Click
End Sub

Private Sub ImgAbout_Click()
LblAbout_Click
End Sub

Private Sub ImgNext_Click()
If Me.TxtCont.Text = Me.TxtContraseña.Text Then Unload FrmLoghon

If Not Me.TxtCont.Text = Me.TxtContraseña.Text Then
Me.TxtContraseña.Text = ""
Me.TxtContraseña.BackColor = &HFFFFFF
End If

End Sub

Private Sub LblAbout_Click()
FrmAbout.Show
End Sub

Private Sub lblLoghon_Click()

ImgLoghon.Left = Screen.Width - Width \ 2
ImgLoghon.Top = Screen.Height - Height \ 2

TxtContraseña.Left = Screen.Width - Width \ 2 - 550
TxtContraseña.Top = Screen.Height - Height \ 2 + 1950

ImgNext.Left = Screen.Width - Width \ 2 + 2350
ImgNext.Top = Screen.Height - Height \ 2 + 1850

ImgLoghon.Visible = True
TxtContraseña.Visible = True
ImgNext.Visible = True
End Sub

Private Sub lblOpciones_Click()
BtnRestart_Click
End Sub

Private Sub LblShutDown_Click()
BtnShutDown_Click
End Sub



Private Sub Tmrwallpaper_Timer()
Me.imgWall.Top = 0
Me.imgWall.Left = 0
Me.imgWall.Width = FrmDesktop.Width
Me.imgWall.Height = FrmDesktop.Height
Tmrwallpaper.Enabled = False

End Sub

Private Sub TxtContraseña_Change()
TxtContraseña.BackColor = &HC0FFC0
End Sub

Private Sub TxtContraseña_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ImgNext_Click
End If
End Sub
