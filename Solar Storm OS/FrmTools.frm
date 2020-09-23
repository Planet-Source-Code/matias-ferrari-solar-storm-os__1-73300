VERSION 5.00
Begin VB.Form FrmTools 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Herramientas"
   ClientHeight    =   6735
   ClientLeft      =   4755
   ClientTop       =   1740
   ClientWidth     =   10080
   Icon            =   "FrmTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   6390
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10050
      TabIndex        =   2
      Top             =   0
      Width           =   10080
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Width           =   8055
      End
      Begin VB.Image BtnGo 
         Height          =   480
         Left            =   9000
         MouseIcon       =   "FrmTools.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "FrmTools.frx":0E1C
         ToolTipText     =   "Ir"
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   9600
         Picture         =   "FrmTools.frx":1AE6
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5880
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   6270
      Left            =   3240
      Pattern         =   "*.exe"
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "FrmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGo_Click()
On Error Resume Next
If txtPath.Text = " Herramientas" Then File1.Path = App.Path & "\Herramientas"
Dir1.Path = txtPath.Text
File1.Path = txtPath.Text
File1.Refresh
File1.Pattern = "*.*"
End Sub

Private Sub Dir1_Change()
txtPath.Text = Dir1.Path
File1.Path = txtPath.Text
File1.Pattern = "*.*"
File1.Refresh

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
Dir1.Path = App.Path & "\Herramientas"
File1.Path = App.Path & "\Herramientas"
Me.txtPath.Text = File1.Path
txtPath.Text = " Herramientas"
txtPath.BackColor = &HC0FFC0
File1.Pattern = "*.*"

End Sub



Private Sub txtPath_Change()
txtPath.BackColor = &HFFFFFF
End Sub
