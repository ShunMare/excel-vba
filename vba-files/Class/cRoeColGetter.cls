VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRoeColGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub SetRowAndColInfo()
    Dim rTarget As Range
    
    If 0 = m_lEndRow And checkMyValueRow And (0 <> m_lKeyCol) Then
        Set rTarget = ws.Cells(ws.Rows.Count, m_lKeyCol).End(xlUp)
        m_lEndRow = rTarget.Row + (rTarget.Rows.MergeArea.Count - 1)
    End If
    If 0 = m_lEndCol And checkMyValueRow And (0 <> m_lKeyRow) Then
        Set rTarget = ws.Cells(m_lKeyRow, ws.Columns.Count).End(xlToLeft)
        m_lEndCol = rTarget.Column + (rTarget.MergeArea.Columns.Count - 1)
    End If
End Sub

