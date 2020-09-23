VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocumentWorpad 
   Caption         =   "frmDocument"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   1995
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDocumentWorpad.frx":0000
   End
End
Attribute VB_Name = "frmDocumentWorpad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rtfText_SelChange()
    fMainForm.tbToolBar.Buttons("Negrita").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Cursiva").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Subrayado").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Alinear a la izquierda").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Centrar").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Alinear a la derecha").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
End Sub

Private Sub Form_Load()
    Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
End Sub

