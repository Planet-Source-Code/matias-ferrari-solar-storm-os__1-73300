VERSION 5.00
Begin VB.Form FrmProcesos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos en Ejecución"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "FrmPorcesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   2760
      Top             =   6120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nueva Tarea"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar Proceso"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   5880
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5685
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refrescar Procesos"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CPU Utilizado : "
      Height          =   255
      Left            =   -120
      TabIndex        =   7
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label LblCPU 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Procesos :"
      Height          =   255
      Left            =   -240
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   5880
      Width           =   855
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare a new instance of the clsCPUUsage class
Private m_clsCPUUsage As New clsCPUUsage

' Variables para usar Wmi
Dim ListaProcesos As Object
Dim ObjetoWMI As Object
Dim ProcesoACerrar As Object


Private Function MatarProceso( _
    StrNombreProceso As String, _
    Optional DecirSINO As Boolean = True) As Boolean

    MatarProceso = False

    Set ObjetoWMI = GetObject("winmgmts:")

    If IsNull(ObjetoWMI) = False Then

    'instanciamos la variable
    Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")

    For Each ProcesoACerrar In ListaProcesos
        If UCase(ProcesoACerrar.Name) = UCase(StrNombreProceso) Then
            If DecirSINO Then
               If MsgBox("¿Matar el proceso " & _
                  ProcesoACerrar.Name & vbNewLine & "...¿Está seguro?", _
                                        vbYesNo + vbCritical) = vbYes Then
                  ProcesoACerrar.Terminate (0)
                  MatarProceso = True
               End If

            Else

            'Matamos el proceso con el método Terminate
                ProcesoACerrar.Terminate (0)
                MatarProceso = True

            End If
        End If

    Next
    End If

    'Elimina las variables
    Set ListaProcesos = Nothing
    Set ObjetoWMI = Nothing
End Function

Private Sub Listar()

    Set ObjetoWMI = GetObject("winmgmts:")

    If IsNull(ObjetoWMI) = False Then
        ' En esta variable se obtienen los procesos
        Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
        'Recorremos toda la coleccion en la lista de procesos _
        y la añadimos al control listbox
        For Each ProcesoACerrar In ListaProcesos
            List1.AddItem LCase$(ProcesoACerrar.Name)
        Next
    End If

    'Eliminamos las variables de objeto

    Set ListaProcesos = Nothing
    Set ObjetoWMI = Nothing

End Sub

Private Sub Command1_Click()

    'Llamamos a MatarProceso pasandole el nombre
    MatarProceso LCase$(List1), True
    'Borramos el list
    List1.Clear
    'Volvemos a listar los procesos
Command2_Click

End Sub

Private Sub Command2_Click()
    'Borramos la lista y volvemos a listar los procesos
    List1.Clear
    Call Listar
    Label1.Caption = List1.ListCount
End Sub

Private Sub Command3_Click()
Unload Me
FrmRun.Show
End Sub

Private Sub Form_Load()
    Command2.Caption = " Refrescar Procesos"
    Command1.Caption = " Cerrar Proceso "
 List1.Clear
    Call Listar
    Label1.Caption = List1.ListCount

End Sub

Private Sub Timer1_Timer()
   LblCPU.Caption = m_clsCPUUsage.CurrentCPUUsage & " %"

End Sub
