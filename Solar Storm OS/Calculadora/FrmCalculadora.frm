VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmCalculadora 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculadora"
   ClientHeight    =   2415
   ClientLeft      =   4140
   ClientTop       =   3810
   ClientWidth     =   3120
   Icon            =   "FrmCalculadora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":08CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "2"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0902
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "3"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":091E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":093A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "5"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0956
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "6"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0972
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "7"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":098E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "8"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":09AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Number 
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "9"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":09C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Decimal 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":09E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Operator 
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "="
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":09FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Operator 
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0A1A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Operator 
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   13
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0A36
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Operator 
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   14
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "/"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0A52
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Operator 
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   15
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "*"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0A6E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Percent 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "%"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0A8A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Cancel 
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "C"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn CancelEntry 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "CE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   16777215
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCalculadora.frx":0AC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
' Autor : Microsoft Corporation

Option Explicit

Dim Op1, Op2                ' Previously input operand.
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput               ' Indicate type of last keypress event.
Dim OpFlag                  ' Indicate pending operation.
Dim TempReadout

' Click event procedure for C (cancel) key.
' Reset the display and initializes variables.
Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

' Click event procedure for CE (cancel entry) key.
Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

' Click event procedure for decimal point (.) key.
' If last keypress was an operator, initialize
' readout to "0." Otherwise, append a decimal
' point to the display.
Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

' Initialization routine for the form.
' Set all variables to initial values.
Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    'Decimal.Caption  = Format(0, ".")
End Sub

Private Sub Image1_Click()

End Sub



' Click event procedure for number keys (0-9).
' Append new number to the number in the display.
Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

' Click event procedure for operator keys (+, -, x, /, =).
' If the immediately preceeding keypress was part of a
' number, increments NumOps. If one operand is present,
' set Op1. If two are present, set Op1 equal to the
' result of the operation on Op1 and the current
' input string, and display the result.
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "*"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Can't divide by zero", 48, "Calculator"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

' Click event procedure for percent key (%).
' Compute and display a percentage of the first operand.
Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub


                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    