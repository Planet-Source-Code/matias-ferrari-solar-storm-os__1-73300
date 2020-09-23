VERSION 5.00
Begin VB.Form FrmSearchandReplaceNotepad 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar y reemplazar"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtwhat 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtreplace 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdfindnext 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdreplace 
      Caption         =   "&Reemplazar"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdreplaceall 
      Caption         =   "Reemplazar todo"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox chkmatchcase 
      Caption         =   "Coincidir May / Min"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form4.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Texto"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   240
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Reemplazar con .."
      Height          =   195
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   1290
   End
End
Attribute VB_Name = "FrmSearchandReplaceNotepad"
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

Private Sub cmdfindnext_Click()

Dim compare As Integer

If chkmatchcase.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If
position = InStr(position + 1, FrmFocus.RichTextBox1.Text, txtwhat.Text, compare)
If position > 0 Then
    FrmFocus.RichTextBox1.SelStart = position - 1
    FrmFocus.SetFocus
Else
    cmdcancel_Click
    MsgBox "No se encontró el texto", vbInformation
End If

End Sub

Private Sub cmdreplace_Click()
Dim compare As Integer

If chkmatchcase.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If
position = InStr(position + 1, FrmFocus.RichTextBox1.Text, txtwhat.Text, compare)
If position > 0 Then
     
    FrmFocus.RichTextBox1.SelStart = position - 1
    FrmFocus.RichTextBox1.SelLength = Len(txtwhat.Text)
    
    FrmFocus.SetFocus
    
    FrmFocus.RichTextBox1.SelText = txtreplace.Text
    
Else
    cmdcancel_Click
    MsgBox "No se encontró el texto", vbInformation
    Exit Sub
End If
cmdcancel_Click
End Sub

Private Sub cmdreplaceall_Click()

Dim compare As Integer
Dim i As Integer
Dim J As Integer


position = 0
i = Len(FrmFocus.RichTextBox1.Text)

If chkmatchcase.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If

For J = 0 To i
position = InStr(position + 1, FrmFocus.RichTextBox1.Text, txtwhat.Text, compare)
If position > 0 Then
    
    FrmFocus.RichTextBox1.SelStart = position - 1
    FrmFocus.RichTextBox1.SelLength = Len(txtwhat.Text)
    FrmFocus.SetFocus
    FrmFocus.RichTextBox1.SelText = txtreplace.Text
  End If
Next
cmdcancel_Click
If position < 0 Then
    MsgBox "No se encontró el texto", vbInformation
End If

End Sub


Private Sub Form_Load()

End Sub
