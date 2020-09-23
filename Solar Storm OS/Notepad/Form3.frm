VERSION 5.00
Begin VB.Form FrmSearchNotepad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtfind 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chkmatch 
      Caption         =   "Coin May / Min"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form3.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Texto"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "FrmSearchNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
position = 0
Unload Me
Set FrmSearchandReplaceNotepad = Nothing
End Sub

Private Sub cmdfind_Click()

Dim compare As Integer

If chkmatch.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If

position = InStr(position + 1, FrmFocus.RichTextBox1.Text, txtfind.Text, compare)

If position > 0 Then
    FrmFocus.RichTextBox1.SelStart = position - 1
    FrmFocus.RichTextBox1.SelLength = Len(txtfind.Text)
    
    FrmFocus.SetFocus
    
Else
    position = 0
    Unload Me
    MsgBox "No se encontró el texto", vbInformation
End If

End Sub
