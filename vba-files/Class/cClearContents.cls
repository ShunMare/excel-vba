VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cClearContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_lStartRow As Long
Dim m_lStartCol As Long
Dim m_lEndCol As Long
Dim m_lKeyCol As Long

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get sWorkbookName() As String
    sWorkbookName = m_sWorkbookName
End Property

Public Property Let sWorkbookName(ByVal vsWorkbookName As String)
    m_sWorkbookName = vsWorkbookName
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get sWorksheetName() As String
    sWorksheetName = m_sWorksheetName
End Property

Public Property Let sWorksheetName(ByVal vsWorksheetName As String)
    m_sWorksheetName = vsWorksheetName
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get lStartRow() As Long
    lStartRow = m_lStartRow
End Property

Public Property Let lStartRow(ByVal vlStartRow As Long)
    m_lStartRow = vlStartRow
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get lKeyCol() As Long
    lKeyCol = m_lKeyCol
End Property

Public Property Let lKeyCol(ByVal vlKeyCol As Long)
    m_lKeyCol = vlKeyCol
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get lStartCol() As Long
    lStartCol = m_lStartCol
End Property

Public Property Let lStartCol(ByVal vlStartCol As Long)
    m_lStartCol = vlStartCol
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get lEndCol() As Long
    lEndCol = m_lEndCol
End Property

Public Property Let lEndCol(ByVal vlEndCol As Long)
    m_lEndCol = vlEndCol
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sWorkbookName = ""
    m_sWorksheetName = ""
End Sub

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Function ClearRangeContents()
    Dim wsTarget As Worksheet
    Dim rTarget As Range
    Dim lEndRow As Long
    
    If sWorkbookName = "" Then
        Set wsTarget = ThisWorkbook.Worksheets(sWorksheetName)
    ElseIf sWorkbookName <> "" Then
        Set wsTarget = Workbooks(sWorkbookName).Worksheets(sWorksheetName)
    Else
        MsgBox "err"
        Exit Function
    End If
    lEndRow = wsTarget.Cells(wsTarget.Rows.Count, lKeyCol).End(xlUp).Row
    
    With wsTarget
        Set rTarget = .Range(.Cells(lStartRow, lStartCol), .Cells(lEndRow, lEndCol))
        rTarget.ClearContents
    End With


End Function
