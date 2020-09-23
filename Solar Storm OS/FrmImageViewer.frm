VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmImageViewer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor de Imagenes"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   Icon            =   "FrmImageViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8535
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   10095
         Begin VB.HScrollBar vBar2 
            Height          =   270
            Left            =   -10
            SmallChange     =   80
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   7080
            Value           =   200
            Width           =   9885
         End
         Begin VB.VScrollBar vBar1 
            Height          =   7005
            Left            =   9840
            SmallChange     =   80
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   80
            Value           =   200
            Width           =   270
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   9840
            Top             =   7080
            Width           =   255
         End
         Begin VB.Image Image1 
            Height          =   7215
            Left            =   0
            Stretch         =   -1  'True
            Top             =   120
            Width           =   9975
         End
      End
      Begin MSComDlg.CommonDialog dlgOpenImg 
         Left            =   7080
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image btnImagen 
         Height          =   720
         Left            =   1440
         MouseIcon       =   "FrmImageViewer.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "FrmImageViewer.frx":0E1C
         ToolTipText     =   "Modo Imagen"
         Top             =   45
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image BtnAllScreen 
         Enabled         =   0   'False
         Height          =   720
         Left            =   1440
         MouseIcon       =   "FrmImageViewer.frx":1CE6
         MousePointer    =   99  'Custom
         Picture         =   "FrmImageViewer.frx":1E38
         ToolTipText     =   "Pantalla Completa"
         Top             =   0
         Width           =   720
      End
      Begin VB.Image optStretch 
         Enabled         =   0   'False
         Height          =   480
         Left            =   240
         MouseIcon       =   "FrmImageViewer.frx":2D02
         MousePointer    =   99  'Custom
         Picture         =   "FrmImageViewer.frx":2E54
         ToolTipText     =   "Ajustar a la Ventana"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image optUnStretch 
         Enabled         =   0   'False
         Height          =   720
         Left            =   720
         MouseIcon       =   "FrmImageViewer.frx":3B1E
         MousePointer    =   99  'Custom
         Picture         =   "FrmImageViewer.frx":3C70
         ToolTipText     =   "Mostrar Tamaño Real"
         Top             =   0
         Width           =   720
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label LblPath 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre : "
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label label0 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta : "
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LblOpen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir Imagen"
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
         Left            =   8760
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image btnOpen 
         Height          =   720
         Left            =   8040
         MouseIcon       =   "FrmImageViewer.frx":4B3A
         MousePointer    =   99  'Custom
         Picture         =   "FrmImageViewer.frx":4C8C
         Top             =   120
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmImageViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnAllScreen_Click()
Me.WindowState = 2

Picture1.Width = Me.Width
Picture1.Height = Me.Height
vBar1.Visible = False
vBar2.Visible = False
Frame1.Left = 0
Frame1.Width = Me.Width
Frame1.Height = Me.Height
Shape1.Visible = False



Image1.Stretch = False
Image1.Top = (Screen.Height - Image1.Height) / 2 - 200
Image1.Left = (Screen.Width - Image1.Width) / 2
btnImagen.Visible = True
BtnAllScreen.Visible = False
End Sub

Private Sub btnImagen_Click()
btnImagen.Visible = False
BtnAllScreen.Visible = True
Me.WindowState = 0
optStretch_Click
vBar1.Visible = True
vBar2.Visible = True
Shape1.Visible = True
Frame1.Left = 120


End Sub

Private Sub btnOpen_Click()
        Image1.Height = "7215"
        Image1.Width = "9975"
        Image1.Top = "100"
        Image1.Left = "22"
    
    With dlgOpenImg
        .DialogTitle = "Abrir Imagen"
        .CancelError = False
        .Filter = "Image files (*.*)|*.*"
        .ShowOpen
        Image1.Picture = LoadPicture(dlgOpenImg.FileName)
        lblName.Caption = .FileTitle
        LblPath.Caption = .FileName
        
    BtnAllScreen.Enabled = True
    optStretch.Enabled = True
    optUnStretch.Enabled = True
        
    End With

End Sub


Private Sub Image1_DblClick()
Me.WindowState = 2

End Sub



Private Sub LblOpen_Click()
btnOpen_Click
End Sub

Private Sub optStretch_Click()
Image1.Stretch = True
        Image1.Height = "7215"
        Image1.Width = "9975"
        Image1.Top = "100"
        Image1.Left = "22"

End Sub

Private Sub optUnStretch_Click()
    Image1.Stretch = False
End Sub

Private Sub Timer1_Timer()
    If Stretch.Text = "True" Then
        Image1.Stretch = True
    End If
    If Stretch.Text = "False" Then
        Image1.Stretch = False
    End If
Timer1.Enabled = True
End Sub



Private Sub vBar1_Change()
    Image1.Top = (-vBar1.Value)

End Sub
Private Sub vBar1_Scroll()
    Call vBar1_Change
End Sub

Private Sub vBar2_Change()
    Image1.Left = (-vBar2.Value)
End Sub

Private Sub vBar2_Scroll()
    Call vBar2_Change
End Sub
