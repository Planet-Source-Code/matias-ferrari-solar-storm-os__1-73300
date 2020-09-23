VERSION 5.00
Begin VB.Form FrmCodigo_Fuente 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Código fuente"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6300
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FrmCodigo_Fuente.frx":0000
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "FrmCodigo_Fuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
Unload Me
End Sub

Private Sub Form_Resize()
Command1.Move Me.ScaleWidth - (Command1.Width + 50), Me.ScaleHeight - (Command1.Height + 50)
Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - (Command1.Height + 100)
End Sub
