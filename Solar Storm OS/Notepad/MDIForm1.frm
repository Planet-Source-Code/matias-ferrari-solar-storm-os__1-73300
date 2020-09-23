VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDINotepad 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12180
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5874
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":614E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7302
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":84B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":966A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A81E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B9D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6915
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   1429
      ButtonWidth     =   1958
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar como"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cortar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copiar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pegar"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   20
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnulinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "Abrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGuardarComo 
         Caption         =   "Guardar Como .."
      End
      Begin VB.Menu mnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuCopiar 
         Caption         =   "Copiar"
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "Cortar"
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "Pegar"
      End
      Begin VB.Menu mnuLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeleccionar 
         Caption         =   "Seleccionar todo"
      End
   End
   Begin VB.Menu mnuBuscarReemplazar 
      Caption         =   "Buscar"
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuLina4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReemplazar 
         Caption         =   "Buscar y reemplazar"
      End
   End
   Begin VB.Menu mnuDocumentos 
      Caption         =   "&Documentos"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de .."
      End
   End
End
Attribute VB_Name = "MDINotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Initialize()
CommonDialog1.Filter = "Documento de texto|*.txt|Todos los Archivos|*.*"
End Sub

'Menú abrir
Private Sub mnuAbrir_Click()

On Error GoTo errSub

CommonDialog1.ShowOpen

If CommonDialog1.FileName <> "" Then
    Set FrmDoc = New frmDocumento
    
    FrmDoc.Show
    
    ActiveForm.Caption = CommonDialog1.FileName

    ActiveForm.RichTextBox1.LoadFile CommonDialog1.FileName

End If
Exit Sub
errSub:

Select Case Err.Number
Case 70
      ActiveForm.RichTextBox1.LoadFile CommonDialog1.FileName
      Resume Next
End Select

End Sub




Private Sub mnuBuscar_Click()
If Forms.Count > 1 Then
    FrmSearchNotepad.Show
    TopMost FrmSearchNotepad
End If
End Sub



Private Sub mnuFuente_Click()

End Sub



'Menu Guardar Como



'Menú para guardar el archivo
Private Sub mnuGuardar_Click()

On Error GoTo errSub

If Forms.Count = 1 Then
   MsgBox "No hay documentos para guardar", vbInformation
   Exit Sub
End If
If InStr(1, ActiveForm.Caption, sCaption) Then
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then Exit Sub
    ActiveForm.RichTextBox1.SaveFile CommonDialog1.FileName
Else
    ActiveForm.RichTextBox1.SaveFile ActiveForm.Caption
    
End If

Exit Sub
errSub:

Select Case Err.Number
  Case 91
     Resume Next
End Select

End Sub

Private Sub mnuGuardarComo_Click()
On Error GoTo errSub

If Forms.Count = 1 Then
   MsgBox "No hay documentos para guardar", vbInformation
   Exit Sub
End If

CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub
ActiveForm.RichTextBox1.SaveFile CommonDialog1.FileName
Exit Sub
errSub:

Select Case Err.Number
  Case 91
    Resume Next
End Select
End Sub

'Menú nuevo archivo
Private Sub mnuNuevo_Click()
Set FrmDoc = New frmDocumento
  nForms = nForms + 1
  FrmDoc.Caption = sCaption & nForms
  FrmDoc.Show
End Sub

'Menú pegar
Private Sub mnuPegar_Click()
On Local Error Resume Next
ActiveForm.RichTextBox1.SelText = Clipboard.GetText
End Sub

Private Sub mnuReemplazar_Click()
If Forms.Count > 1 Then
    FrmSearchandReplaceNotepad.Show
    TopMost FrmSearchandReplaceNotepad
End If

End Sub

'Menú salir
Private Sub mnuSalir_Click()
    Unload Me
End Sub

'Menu para seleccionar todo el texto
Private Sub mnuSeleccionar_Click()
On Local Error Resume Next

ActiveForm.RichTextBox1.SelStart = 0
ActiveForm.RichTextBox1.SelLength = Len(ActiveForm.RichTextBox1.Text)

End Sub



'Menú Copiar texto
Private Sub mnuCopiar_Click()
On Local Error Resume Next
Clipboard.SetText ActiveForm.RichTextBox1.SelText
End Sub

'Menú cortar texto
Private Sub mnuCortar_Click()
On Local Error Resume Next
Clipboard.SetText ActiveForm.RichTextBox1.SelText
ActiveForm.RichTextBox1.SelText = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)


    Select Case Button.Index
    
        Case 1: mnuNuevo_Click
        Case 3: mnuAbrir_Click
        Case 5: mnuGuardar_Click
        Case 6: mnuGuardarComo_Click
        Case 8: mnuCortar_Click
        Case 9: mnuCopiar_Click
        Case 10: mnuPegar_Click
        Case 12: mnuBuscar_Click
    
    End Select
    

End Sub
