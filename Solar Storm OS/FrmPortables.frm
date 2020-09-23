VERSION 5.00
Begin VB.Form FrmPortables 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Portables"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "FrmPortables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   6465
      Left            =   3240
      Pattern         =   "*.exe"
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   3240
      ScaleHeight     =   6345
      ScaleWidth      =   6945
      TabIndex        =   6
      Top             =   480
      Width           =   6975
      Begin VB.Label lblGames 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Juegos"
         Height          =   255
         Left            =   2400
         MouseIcon       =   "FrmPortables.frx":0CCA
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.Image ImgGames 
         Height          =   720
         Left            =   2520
         MouseIcon       =   "FrmPortables.frx":0E1C
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":0F6E
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image ImgDiseño 
         Height          =   720
         Left            =   2400
         MouseIcon       =   "FrmPortables.frx":2C38
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":2D8A
         Top             =   120
         Width           =   720
      End
      Begin VB.Label LblDiseño 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diseño Grafico"
         Height          =   255
         Left            =   2280
         MouseIcon       =   "FrmPortables.frx":4A54
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image ImgOfimatica 
         Height          =   720
         Left            =   6000
         MouseIcon       =   "FrmPortables.frx":4BA6
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":4CF8
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image ImgReproductores 
         Height          =   720
         Left            =   360
         MouseIcon       =   "FrmPortables.frx":69C2
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":6B14
         Top             =   2880
         Width           =   720
      End
      Begin VB.Image ImgUtilities 
         Height          =   720
         Left            =   1560
         MouseIcon       =   "FrmPortables.frx":87DE
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":8930
         Top             =   2880
         Width           =   720
      End
      Begin VB.Image ImgGrabacion 
         Height          =   720
         Left            =   240
         MouseIcon       =   "FrmPortables.frx":A5FA
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":A74C
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image ImgInternet 
         Height          =   720
         Left            =   1440
         MouseIcon       =   "FrmPortables.frx":C416
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":C568
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image ImgMensajeria 
         Height          =   720
         Left            =   3600
         MouseIcon       =   "FrmPortables.frx":E232
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":E384
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image ImgNavegadores 
         Height          =   720
         Left            =   4800
         MouseIcon       =   "FrmPortables.frx":1004E
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":101A0
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image ImgGestores 
         Height          =   720
         Left            =   6000
         MouseIcon       =   "FrmPortables.frx":11E6A
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":11FBC
         Top             =   120
         Width           =   720
      End
      Begin VB.Label LblUtilities 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Utilidades"
         Height          =   255
         Left            =   1440
         MouseIcon       =   "FrmPortables.frx":13C86
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblReproductores 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reprodctores Multimedia"
         Height          =   495
         Left            =   120
         MouseIcon       =   "FrmPortables.frx":13DD8
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblOfimatica 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ofimatica"
         Height          =   255
         Left            =   5880
         MouseIcon       =   "FrmPortables.frx":13F2A
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LblNavegadores 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Navegadores"
         Height          =   255
         Left            =   4680
         MouseIcon       =   "FrmPortables.frx":1407C
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LblMensajeria 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mensajeria"
         Height          =   255
         Left            =   3480
         MouseIcon       =   "FrmPortables.frx":141CE
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LblInternet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Internet"
         Height          =   255
         Left            =   1320
         MouseIcon       =   "FrmPortables.frx":14320
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LblGrabacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Grabación"
         Height          =   255
         Left            =   120
         MouseIcon       =   "FrmPortables.frx":14472
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblGestores 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gestores de Descarga"
         Height          =   495
         Left            =   5760
         MouseIcon       =   "FrmPortables.frx":145C4
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblMail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electronico"
         Height          =   495
         Left            =   4680
         MouseIcon       =   "FrmPortables.frx":14716
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LblFTP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clientes FTP"
         Height          =   255
         Left            =   3480
         MouseIcon       =   "FrmPortables.frx":14868
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LblAudio 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Audio"
         Height          =   255
         Left            =   1200
         MouseIcon       =   "FrmPortables.frx":149BA
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Image ImgAntivirus 
         Height          =   720
         Left            =   240
         MouseIcon       =   "FrmPortables.frx":14B0C
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":14C5E
         Top             =   120
         Width           =   720
      End
      Begin VB.Image ImgAudio 
         Height          =   720
         Left            =   1320
         MouseIcon       =   "FrmPortables.frx":16928
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":16A7A
         Top             =   120
         Width           =   720
      End
      Begin VB.Image ImgFTP 
         Height          =   720
         Left            =   3600
         MouseIcon       =   "FrmPortables.frx":18744
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":18896
         Top             =   120
         Width           =   720
      End
      Begin VB.Image ImgMail 
         Height          =   720
         Left            =   4800
         MouseIcon       =   "FrmPortables.frx":1A560
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":1A6B2
         Top             =   120
         Width           =   720
      End
      Begin VB.Label LblAntivirus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Antivirus"
         Height          =   255
         Left            =   120
         MouseIcon       =   "FrmPortables.frx":1C37C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10170
      TabIndex        =   1
      Top             =   0
      Width           =   10200
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   3240
         ScaleHeight     =   15
         ScaleWidth      =   5295
         TabIndex        =   5
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   8055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   9600
         Picture         =   "FrmPortables.frx":1C4CE
         Top             =   0
         Width           =   480
      End
      Begin VB.Image BtnGo 
         Height          =   480
         Left            =   9000
         MouseIcon       =   "FrmPortables.frx":1D198
         MousePointer    =   99  'Custom
         Picture         =   "FrmPortables.frx":1D2EA
         ToolTipText     =   "Ir"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   6390
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "FrmPortables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGo_Click()
On Error Resume Next
If txtPath.Text = " Portables" Then File1.Path = App.Path & "\Portables"
Dir1.Path = txtPath.Text
File1.Path = txtPath.Text
File1.Refresh
File1.Pattern = "*.*"
File1.Visible = True

End Sub

Private Sub Dir1_Change()
txtPath.Text = Dir1.Path
File1.Path = txtPath.Text
File1.Pattern = "*.*"
File1.Refresh
File1.Visible = True
If File1.Path = App.Path & "\Portables" Then File1.Visible = False

End Sub

Private Sub File1_Click()
Text1.Text = File1.Path + "\" + File1.FileName
End Sub

Private Sub File1_DblClick()
On Error Resume Next
Shell (Text1.Text), vbNormalFocus
Unload Me

End Sub

Private Sub Form_Load()
Dir1.Path = App.Path & "\Portables"
File1.Path = App.Path & "\Portables"
Me.txtPath.Text = File1.Path
txtPath.Text = " Portables"
txtPath.BackColor = &HC0FFC0
End Sub





Private Sub ImgAntivirus_Click()
LblAntivirus_Click
End Sub



Private Sub ImgAudio_Click()
LblAudio_Click
End Sub

Private Sub ImgDiseño_Click()
LblDiseño_Click
End Sub

Private Sub ImgFTP_Click()
LblFTP_Click
End Sub

Private Sub ImgGames_Click()
lblGames_Click
End Sub

Private Sub ImgGestores_Click()
lblGestores_Click
End Sub

Private Sub ImgGrabacion_Click()
LblGrabacion_Click
End Sub

Private Sub ImgInternet_Click()
LblInternet_Click
End Sub

Private Sub ImgMail_Click()
lblMail_Click
End Sub

Private Sub ImgMensajeria_Click()
LblMensajeria_Click
End Sub

Private Sub ImgNavegadores_Click()
LblNavegadores_Click
End Sub

Private Sub ImgOfimatica_Click()
lblOfimatica_Click
End Sub

Private Sub ImgReproductores_Click()
lblReproductores_Click
End Sub

Private Sub ImgUtilities_Click()
LblUtilities_Click
End Sub

Private Sub LblAntivirus_Click()
Dir1.Path = App.Path & "\Portables\Antivirus"
Me.txtPath.Text = File1.Path
txtPath.Text = " Antivirus"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub LblAudio_Click()
Dir1.Path = App.Path & "\Portables\Audio"
Me.txtPath.Text = File1.Path
txtPath.Text = " Audio"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub LblDiseño_Click()
Dir1.Path = App.Path & "\Portables\Diseño Grafico"
Me.txtPath.Text = File1.Path
txtPath.Text = " Diseño Grafico"
txtPath.BackColor = &HC0FFC0
File1.Visible = True
End Sub

Private Sub LblFTP_Click()
Dir1.Path = App.Path & "\Portables\Clientes FTP"
Me.txtPath.Text = File1.Path
txtPath.Text = " Clientes FTP"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub lblGames_Click()
Dir1.Path = App.Path & "\Games"
Me.txtPath.Text = File1.Path
txtPath.Text = " Internet"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub lblGestores_Click()
Dir1.Path = App.Path & "\Portables\Gestores de Descarga"
Me.txtPath.Text = File1.Path
txtPath.Text = " Gestores de Descarga"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub LblGrabacion_Click()
Dir1.Path = App.Path & "\Portables\Grabacion"
Me.txtPath.Text = File1.Path
txtPath.Text = " Grabacion"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub LblInternet_Click()
Dir1.Path = App.Path & "\Portables\Internet"
Me.txtPath.Text = File1.Path
txtPath.Text = " Internet"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub lblMail_Click()
Dir1.Path = App.Path & "\Portables\Correo Electronico"
Me.txtPath.Text = File1.Path
txtPath.Text = " Correo Electronico"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub LblMensajeria_Click()
Dir1.Path = App.Path & "\Portables\Mensajeria Instantanea"
Me.txtPath.Text = File1.Path
txtPath.Text = " Mensajeria Instantanea"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub LblNavegadores_Click()
Dir1.Path = App.Path & "\Portables\Navegadores"
Me.txtPath.Text = File1.Path
txtPath.Text = " Navegadores"
txtPath.BackColor = &HC0FFC0
File1.Visible = True
End Sub

Private Sub lblOfimatica_Click()
Dir1.Path = App.Path & "\Portables\Ofimatica"
Me.txtPath.Text = File1.Path
txtPath.Text = " Ofimatica"
txtPath.BackColor = &HC0FFC0
File1.Visible = True
End Sub

Private Sub lblReproductores_Click()
Dir1.Path = App.Path & "\Portables\Reproductores Multimedia"
Me.txtPath.Text = File1.Path
txtPath.Text = " Reproductores Multimedia"
txtPath.BackColor = &HC0FFC0
File1.Visible = True
End Sub

Private Sub LblUtilities_Click()
Dir1.Path = App.Path & "\Portables\Utilidades"
Me.txtPath.Text = File1.Path
txtPath.Text = " Utilidades"
txtPath.BackColor = &HC0FFC0
File1.Visible = True

End Sub

Private Sub txtPath_Change()
txtPath.BackColor = &HFFFFFF
End Sub

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               