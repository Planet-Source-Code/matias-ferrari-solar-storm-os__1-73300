VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocumento 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   7245
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2990
      _Version        =   393217
      TextRTF         =   $"frmDocumento.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Public flagGuardar As Boolean


Private Sub Form_Resize()
'Redimensionamos el control RichtextBox al ancho y alto del formulario
RichTextBox1.Move ScaleLeft, ScaleTop, ScaleWidth, ScaleHeight
If WindowState = vbMaximized Then
   'mdiform1.Caption = Me.Caption
Else
   MDINotepad.Caption = ""
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

On Error GoTo errSub

Dim ret As Integer

If flagGuardar Then
    ret = MsgBox("¿ Guardar cambios ?", vbYesNoCancel + vbQuestion)
End If

Select Case ret
    Case vbYes:
    
        If InStr(1, Me.Caption, sCaption) Then
            CommonDialog1.ShowSave
            RichTextBox1.SaveFile CommonDialog1.FileName
        Else
            RichTextBox1.SaveFile Me.Caption
        End If
    Case vbCancel:
         Exit Sub
End Select

Set FrmDoc = Nothing
Exit Sub
errSub:

Select Case Err.Number
  Case 75
   Resume Next
   
End Select

End Sub

Private Sub RichTextBox1_Change()
flagGuardar = True
End Sub
            
Private Sub RichTextBox1_GotFocus()
Set FrmFocus = Me
End Sub
