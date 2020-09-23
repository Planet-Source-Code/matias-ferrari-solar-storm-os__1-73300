VERSION 5.00
Begin VB.Form FrmConfiguracion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "FrmConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   9045
      TabIndex        =   7
      Top             =   0
      Width           =   9075
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   3240
         ScaleHeight     =   15
         ScaleWidth      =   5295
         TabIndex        =   9
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Text            =   "Configuración"
         Top             =   120
         Width           =   8055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   9840
         Picture         =   "FrmConfiguracion.frx":0CCA
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Image imgModem 
      Height          =   480
      Left            =   8040
      MouseIcon       =   "FrmConfiguracion.frx":1994
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":1AE6
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modem"
      Height          =   255
      Left            =   7920
      MouseIcon       =   "FrmConfiguracion.frx":27B0
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgEnergia 
      Height          =   480
      Left            =   7080
      MouseIcon       =   "FrmConfiguracion.frx":2902
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":2A54
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Energia"
      Height          =   255
      Left            =   6960
      MouseIcon       =   "FrmConfiguracion.frx":371E
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgJoystick 
      Height          =   480
      Left            =   6120
      MouseIcon       =   "FrmConfiguracion.frx":3870
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":39C2
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Joystick"
      Height          =   255
      Left            =   6000
      MouseIcon       =   "FrmConfiguracion.frx":468C
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgFecha 
      Height          =   480
      Left            =   5040
      MouseIcon       =   "FrmConfiguracion.frx":47DE
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":4930
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha y Hora"
      Height          =   375
      Left            =   4920
      MouseIcon       =   "FrmConfiguracion.frx":55FA
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image imgPantalla 
      Height          =   480
      Left            =   3960
      MouseIcon       =   "FrmConfiguracion.frx":574C
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":589E
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pantalla"
      Height          =   255
      Left            =   3840
      MouseIcon       =   "FrmConfiguracion.frx":6568
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image ImgFirewall 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "FrmConfiguracion.frx":66BA
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":680C
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Firewall"
      Height          =   255
      Left            =   2760
      MouseIcon       =   "FrmConfiguracion.frx":74D6
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image ImgSonido 
      Height          =   480
      Left            =   1800
      MouseIcon       =   "FrmConfiguracion.frx":7628
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":777A
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sonido"
      Height          =   255
      Left            =   1680
      MouseIcon       =   "FrmConfiguracion.frx":8444
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4680
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   360
      X2              =   6960
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema Operativo."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Solar Storm OS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   360
      X2              =   6960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label LblSistema 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema"
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmConfiguracion.frx":8596
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4680
      Width           =   975
   End
   Begin VB.Image imgSistema 
      Height          =   720
      Left            =   480
      MouseIcon       =   "FrmConfiguracion.frx":86E8
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":883A
      Top             =   3960
      Width           =   720
   End
   Begin VB.Label LblActualizar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar"
      Height          =   255
      Left            =   1440
      MouseIcon       =   "FrmConfiguracion.frx":A504
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image ImgActualizar 
      Height          =   480
      Left            =   1680
      MouseIcon       =   "FrmConfiguracion.frx":A656
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":A7A8
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label LblWallpaper 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wallpaper"
      Height          =   255
      Left            =   3720
      MouseIcon       =   "FrmConfiguracion.frx":B472
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image ImgWallpaper 
      Height          =   720
      Left            =   3840
      MouseIcon       =   "FrmConfiguracion.frx":B5C4
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":B716
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label LblAcercaDe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Acerca De..."
      Height          =   255
      Left            =   360
      MouseIcon       =   "FrmConfiguracion.frx":C5E0
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image ImgAcercaDe 
      Height          =   480
      Left            =   600
      MouseIcon       =   "FrmConfiguracion.frx":C732
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":C884
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label LblUsuario 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios"
      Height          =   255
      Left            =   2640
      MouseIcon       =   "FrmConfiguracion.frx":D54E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image ImgUsuario 
      Height          =   720
      Left            =   2760
      MouseIcon       =   "FrmConfiguracion.frx":D6A0
      MousePointer    =   99  'Custom
      Picture         =   "FrmConfiguracion.frx":D7F2
      Top             =   1320
      Width           =   720
   End
End
Attribute VB_Name = "FrmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














Private Sub Image1_Click()

End Sub

Private Sub ImgAcercaDe_Click()
FrmAbout.Show
End Sub

Private Sub ImgActualizar_Click()
Dim ie As Object

Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate "http://www.hunterdesigns.tk"

End Sub

Private Sub imgEnergia_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL powercfg.cpl,,0"
End Sub

Private Sub imgFecha_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL timedate.cpl,,0"
End Sub

Private Sub ImgFirewall_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL firewall.cpl,,0"

End Sub

Private Sub imgJoystick_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL joy.cpl,,0"
End Sub

Private Sub imgModem_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL telephon.cpl,,0"
End Sub

Private Sub imgPantalla_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL desk.cpl,,0"
End Sub

Private Sub imgSistema_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL sysdm.cpl,,0"

End Sub

Private Sub ImgSonido_Click()
Shell "rundll32.exe shell32.dll, Control_RunDLL mmsys.cpl,,0"

End Sub

Private Sub ImgUsuario_Click()
FrmUsuario.Show
End Sub

Private Sub ImgWallpaper_Click()
FrmWallpapers.Show
End Sub

Private Sub Label1_Click()
ImgSonido_Click
End Sub

Private Sub Label10_Click()
imgEnergia_Click
End Sub

Private Sub Label2_Click()
ImgFirewall_Click
End Sub

Private Sub Label3_Click()
imgPantalla_Click
End Sub

Private Sub Label5_Click()
imgFecha_Click
End Sub

Private Sub Label6_Click()
imgJoystick_Click
End Sub

Private Sub Label9_Click()
imgEnergia_Click
End Sub

Private Sub LblAcercaDe_Click()
ImgAcercaDe_Click
End Sub

Private Sub LblActualizar_Click()
ImgActualizar_Click
End Sub

Private Sub LblSistema_Click()
imgSistema_Click
End Sub

Private Sub LblUsuario_Click()
ImgUsuario_Click
End Sub

Private Sub LblWallpaper_Click()
ImgWallpaper_Click
End Sub
