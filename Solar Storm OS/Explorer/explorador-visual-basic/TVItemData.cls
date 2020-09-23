VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTVItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Copyright © 1997-1999 Brad Martinez, http://www.mvps.org
'
Private m_pidlFQ As Long
Private m_pidlRel As Long
'

Public Property Get pidlFQ() As Long
  pidlFQ = m_pidlFQ
End Property

Public Property Let pidlFQ(pidl As Long)
  m_pidlFQ = pidl
End Property

Public Property Get pidlRel() As Long
  pidlRel = m_pidlRel
End Property

Public Property Let pidlRel(pidl As Long)
  m_pidlRel = pidl
End Property
