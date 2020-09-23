VERSION 5.00
Begin VB.Form FrmSplashBrowser 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1850
      Left            =   3000
      Top             =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """ ... will make you internet connection faster better and stronger! """
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
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   6495
   End
   Begin VB.Label LblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Versio!"
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version : "
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2475
      Left            =   0
      Picture         =   "FrmSplashBowser.frx":0000
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   810
      Left            =   3120
      Picture         =   "FrmSplashBowser.frx":182FA
      Top             =   120
      Width           =   3300
   End
End
Attribute VB_Name = "FrmSplashBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Timer1_Timer()

frmPrincipalBrowser.Show
Timer1.Enabled = False
Unload Me
End Sub
