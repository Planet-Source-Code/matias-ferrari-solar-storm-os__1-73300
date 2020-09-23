VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRun 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ejecutar"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ejecutar"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Abrir :"
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
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba el Nombre del programa, documento, carpeta o recurso que desea abrir. Tambien puede buscarlo desde esta ventana."
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FrmEjecutar.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "FrmRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.Filter = "All Files|*.*"
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
Text1.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
Shell (Text1.Text), vbNormalFocus
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
FrmRun.Top = FrmDesktop.PicToolbar.Top - 2340
FrmRun.Left = 450
FrmDesktop.TmrDesktop.Enabled = True
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2_Click
End If
End Sub
