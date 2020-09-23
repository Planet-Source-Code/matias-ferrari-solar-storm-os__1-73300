VERSION 5.00
Begin VB.Form FrmPropiedadesdeUnidad 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades de la Unidad"
   ClientHeight    =   7980
   ClientLeft      =   8100
   ClientTop       =   2280
   ClientWidth     =   7200
   Icon            =   "FrmPropiedadesdeUnidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "General :"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   570
         Width           =   5655
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   720
         ScaleHeight     =   3615
         ScaleWidth      =   5175
         TabIndex        =   2
         Top             =   3960
         Width           =   5175
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   4920
         Top             =   5760
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   5160
         TabIndex        =   1
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   19
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Sin Determinar"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label LblTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad  Conectada"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000000000000000"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000000000000000"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00000000000000000"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   6840
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   6840
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de Archivos :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6720
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "FrmPropiedadesdeUnidad.frx":038A
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espacio libre :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   8
         Top             =   2880
         Width           =   1230
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   360
         Top             =   2520
         Width           =   255
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   360
         Top             =   2880
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   360
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   7
         Top             =   3600
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Espacio utilizado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblCapacidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         Top             =   3600
         Width           =   645
      End
      Begin VB.Label lblUtilizado 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
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
         Height          =   195
         Left            =   5520
         TabIndex        =   4
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label lblLibre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5520
         TabIndex        =   3
         Top             =   2880
         Width           =   645
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Label4"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   3720
      Width           =   2895
   End
End
Attribute VB_Name = "FrmPropiedadesdeUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Drive1_Change()
    Dim tDrive As T_Info
    ' obtener info
    tDrive = getInfoDrive(Left(Drive1.Drive, 3))
    
    ' dibujar
    With tDrive
        
        Call Dibujar_Circulo( _
            .CapacidadBytes, _
            .LibreBytes, 110, _
            &HE0E0E0, vbGreen, _
            RGB(52, 114, 197), vbBlack, Picture1)
        
        ' captions
        lblCapacidad.Caption = .Capacidad
        lblLibre.Caption = .Libre
        lblUtilizado.Caption = .Usado
    End With

End Sub






Private Sub Form_Unload(Cancel As Integer)
FrmDesktop.FrmUnidades.Visible = False
End Sub

Private Sub lblCapacidad_Change()
Label7.Caption = Val(Me.lblCapacidad) * 1024 * 1024 * 1024 * 1024

End Sub

Private Sub lblLibre_Change()
Label5.Caption = Val(Me.lblLibre) * 1024 * 1024 * 1024 * 1024
End Sub

Private Sub lblUtilizado_Change()
Label4.Caption = Val(Me.lblUtilizado) * 1024 * 1024 * 1024 * 1024
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Drive1_Change
End Sub
