VERSION 5.00
Begin VB.Form FrmUsuario 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario"
   ClientHeight    =   4050
   ClientLeft      =   4725
   ClientTop       =   1950
   ClientWidth     =   6735
   Icon            =   "FrmUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6735
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   850
      Left            =   1680
      Top             =   3480
   End
   Begin VB.TextBox TxtCont 
      DataField       =   "Contraseña"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   0
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
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Contraseña"
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2880
      Width           =   6255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2280
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña Cambiada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Escribir Nuevamente la Contraseña :"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nueva Contraseña :"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Escribir contraseña Anterior :"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Crear una Contraseña para su Cuenta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   360
      Picture         =   "FrmUsuario.frx":0CCA
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FrmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()


If Not Text3.Text = Text2.Text Then
MsgBox "Error, Revise las Contraseñas."

Else

If Text1.Text = TxtCont.Text Then
TxtCont.Text = Text3.Text
Label5.Visible = True
Me.Timer1.Enabled = True
Else
MsgBox "Error, Revise las Contraseñas."

End If
End If

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\BD.ss"

End Sub

Private Sub Text1_Change()
Text1.BackColor = &HC0FFC0
Text2.BackColor = vbWhite
Text3.BackColor = vbWhite
End Sub

Private Sub Text2_Change()
Text2.BackColor = &HC0FFC0
Text1.BackColor = vbWhite
Text3.BackColor = vbWhite
End Sub

Private Sub Text3_Change()
Text3.BackColor = &HC0FFC0
Text2.BackColor = vbWhite
Text1.BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
