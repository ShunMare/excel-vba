VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRowColInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lStartRow As Long
Dim m_lStartCol As Long
Dim m_lEndRow As Long
Dim m_lEndCol As Long

Public Property Get lStartRow() As Long
    lStartRow = m_lStartRow
End Property

Public Property Let lStartRow(ByVal vlStartRow As Long)
    m_lStartRow = vlStartRow
End Property

Public Property Get lStartCol() As Long
    lStartCol = m_lStartCol
End Property

Public Property Let lStartCol(ByVal vlStartCol As Long)
    m_lStartCol = vlStartCol
End Property

Public Property Get lEndRow() As Long
    lEndRow = m_lEndRow
End Property

Public Property Let lEndRow(ByVal vlEndRow As Long)
    m_lEndRow = vlEndRow
End Property

Public Property Get lEndCol() As Long
    lEndCol = m_lEndCol
End Property

Public Property Let lEndCol(ByVal vlEndCol As Long)
    m_lEndCol = vlEndCol
End Property

Private Sub Class_Initialize()
    m_lStartRow = 0
    m_lStartCol = 0
    m_lEndRow = 0
    m_lEndCol = 0
End Sub

Private Function checkMyValue()
    checkMyValue = True
End Function
