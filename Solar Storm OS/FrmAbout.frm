VERSION 5.00
Begin VB.Form FrmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Solar Storm"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   4710
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   5280
      Left            =   0
      Picture         =   "FrmAbout.frx":0000
      Top             =   0
      Width           =   8115
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
