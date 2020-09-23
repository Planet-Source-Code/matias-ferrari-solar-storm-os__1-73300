Attribute VB_Name = "Galleta"
Option Explicit

Const PI = 3.14159

Const Kb As Double = 1024
Const Mb As Double = 1024 * Kb
Const Gb As Double = 1024 * Mb
Const Tb As Double = 1024 * Gb
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" ( _
    ByVal lpRootPathName As String, _
    lpFreeBytesAvailableToCaller As LARGE_INTEGER, _
    lpTotalNumberOfBytes As LARGE_INTEGER, _
    lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Type T_Info

    Capacidad As String
    Libre As String
    Usado As String
    
    CapacidadBytes As Double
    LibreBytes     As Double
    UsadoBytes     As Double
End Type

Private Function Entero_a_Double(l As Long, h As Long) As Double

Dim ret As Double

    ret = h
    If h < 0 Then ret = ret + 2 ^ 32
    ret = ret * 2 ^ 32

    ret = ret + l
    If l < 0 Then ret = ret + 2 ^ 32

    Entero_a_Double = ret
End Function

Private Function Size(ByVal n_bytes As Double) As String

    If n_bytes < Kb Then
        Size = Format$(n_bytes) & " bytes"
    ElseIf n_bytes < Mb Then
        Size = Format$(n_bytes / Kb, "0.00") & " KB"
    ElseIf n_bytes < Gb Then
        Size = Format$(n_bytes / Mb, "0.00") & " MB"
    Else
        Size = Format$(n_bytes / Gb, "0.00") & " GB"
    End If
End Function

Function getInfoDrive(Drive As String) As T_Info

On Error GoTo errSub
    
    Dim Avalables As LARGE_INTEGER, Total As LARGE_INTEGER
    Dim Libres As LARGE_INTEGER, dTotal As Double, dLibre As Double
    Dim ret As Long
    
    ret = GetDiskFreeSpaceEx(Drive, Avalables, Total, Libres)
    
    dTotal = Entero_a_Double(Total.lowpart, Total.highpart)
    dLibre = Entero_a_Double(Libres.lowpart, Libres.highpart)

    ' retorna a la función los valores convertidos a String
    With getInfoDrive
        
        ' bytes
        .CapacidadBytes = dTotal
        .LibreBytes = dLibre
        
        ' string
        .Capacidad = Size(dTotal)
        .Libre = Size(dLibre)
        .Usado = Size(dTotal - dLibre)
    End With
    
Exit Function

'Error
errSub:
MsgBox Err.Description, vbCritical

End Function

Sub Dibujar_Circulo( _
    Valor_Maximo As Double, _
    Valor As Double, _
    Radio As Integer, _
    BackColor As Long, _
    ForeColor As Long, _
    ValueColor As Long, _
    BorderColor As Long, _
    Objeto As Object)
    
    If Valor_Maximo <= 0 Then Exit Sub
    
    With Objeto
        .BackColor = BackColor
        .ScaleMode = vbPixels
        Objeto.Cls

        Dim I As Long, per, xs, ys, cx, cy

        per = Valor / Valor_Maximo * 100
        per = per / 100
        per = 360 * per
    
        cx = .ScaleWidth \ 2
        cy = .ScaleHeight \ 2

        .DrawWidth = 2
    End With

    For I = 0 To 360
        xs = Cos(I / 180 * PI) * Radio
        ys = Sin(I / 180 * PI) * Radio
        Objeto.Line (cx, cy)-(cx + xs, cy + ys), ForeColor
        DoEvents
    Next I

    For I = 0 To per
        xs = Cos(I / 180 * PI) * Radio
        ys = Sin(I / 180 * PI) * Radio
        Objeto.Line (cx, cy)-(cx + xs, cy + ys), ValueColor
        DoEvents
    Next I
    
    With Objeto
        .DrawWidth = 2
        Objeto.Circle (.ScaleWidth / 2, .ScaleHeight / 2), Radio + 4, BorderColor
    End With
    
End Sub



