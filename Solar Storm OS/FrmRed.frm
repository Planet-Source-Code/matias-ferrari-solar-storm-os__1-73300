VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmRed 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Propiedades de Red"
   ClientHeight    =   3135
   ClientLeft      =   20745
   ClientTop       =   15060
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Soporte : "
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   4815
      Begin VB.Label LblIP 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lblTipoIP 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección IP :"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de IP :"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conexión : "
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2040
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Refrescar"
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
         Left            =   3720
         MouseIcon       =   "FrmRed.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label LblVelocidad 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label LblDuración 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LblEstado 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duración :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Image ImgRedOff 
         Height          =   960
         Left            =   3720
         Picture         =   "FrmRed.frx":0152
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image ImgRedON 
         Height          =   960
         Left            =   3720
         Picture         =   "FrmRed.frx":339C
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image ImgRed 
         Height          =   855
         Left            =   3720
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmRed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Función Api IsNetworkAlive para detectar _
 si estamos conectados y a que tipo de red
Private Declare Function IsNetworkAlive Lib "SENSAPI.DLL" ( _
    ByRef lpdwFlags As Long) As Long





Private Sub Form_Load()
 'Si la Api retorna 0 quiere decir que no hay ningun tipo de conexión de Red
        If IsNetworkAlive(ret) = 0 Then

            ImgRed.Picture = ImgRedOff.Picture
            ImgRed.ToolTipText = "Conexion a Internet OFF"
            Me.LblDuración.Caption = "00:00:00"
            Me.LblEstado.Caption = "Desconectado"
            Me.LblVelocidad.Caption = "Conectividad Nula o Limitada"
            Me.lblTipoIP.Caption = "No se reconoce o no Asignada"
            Me.LblIP.Caption = "127.0.0.1"
        
        Else
            ' hay conexión , y muestra el tipo
            ImgRed.Picture = ImgRedON.Picture
            ImgRed.ToolTipText = "Conexion a Internet ON"
            Me.LblDuración.Caption = "00:00:00"
            Me.LblEstado.Caption = "Conectado"
            Me.LblVelocidad.Caption = "100mbps"
            Me.lblTipoIP.Caption = "Asignada por DHCP"
            Me.LblIP.Caption = Me.Winsock1.LocalIP
  
  
    End If

End Sub


Private Sub Label6_Click()
            If IsNetworkAlive(ret) = 0 Then

            ImgRed.Picture = ImgRedOff.Picture
            ImgRed.ToolTipText = "Conexion a Internet OFF"
            Me.LblDuración.Caption = "00:00:00"
            Me.LblEstado.Caption = "Desconectado"
            Me.LblVelocidad.Caption = "Conectividad Nula o Limitada"
            Me.lblTipoIP.Caption = "No se reconoce o no Asignada"
            Me.LblIP.Caption = "127.0.0.1"
        
        Else
            ' hay conexión , y muestra el tipo
            ImgRed.Picture = ImgRedON.Picture
            ImgRed.ToolTipText = "Conexion a Internet ON"
            Me.LblDuración.Caption = "00:00:00"
            Me.LblEstado.Caption = "Conectado"
            Me.LblVelocidad.Caption = "100mbps"
            Me.lblTipoIP.Caption = "Asignada por DHCP"
            Me.LblIP.Caption = Me.Winsock1.LocalIP
  
  
    End If

End Sub
