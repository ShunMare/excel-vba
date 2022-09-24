VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_wb As Workbook
Dim m_ws As Worksheet
Dim m_lStartRow As Long
Dim m_lStartCol As Long
Dim m_lEndRow As Long
Dim m_lEndCol As Long
Dim m_lKeyRow As Long
Dim m_lKeyCol As Long
Dim m_lTargetRow As Long
Dim m_lTargetCol As Long
Dim m_lResultRow As Long
Dim m_lResultCol As Long

'--------------------------------------------------------------------------------
'@brief : getter letter sWorkbookName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sWorkbookName() As String
    sWorkbookName = m_sWorkbookName
End Property

Public Property Let sWorkbookName(ByVal vsWorkbookName As String)
    m_sWorkbookName = vsWorkbookName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sWorksheetName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sWorksheetName() As String
    sWorksheetName = m_sWorksheetName
End Property

Public Property Let sWorksheetName(ByVal vsWorksheetName As String)
    m_sWorksheetName = vsWorksheetName
End Property

'--------------------------------------------------------------------------------
'@brief : getter wb
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get wb() As Workbook
    Set wb = m_wb
End Property

'--------------------------------------------------------------------------------
'@brief : getter ws
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get ws() As Worksheet
    Set ws = m_ws
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lStartRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lStartRow() As Long
    lStartRow = m_lStartRow
End Property

Public Property Let lStartRow(ByVal vlStartRow As Long)
    m_lStartRow = vlStartRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lStartCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lStartCol() As Long
    lStartCol = m_lStartCol
End Property

Public Property Let lStartCol(ByVal vlStartCol As Long)
    m_lStartCol = vlStartCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lEndRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lEndRow() As Long
    lEndRow = m_lEndRow
End Property

Public Property Let lEndRow(ByVal vlEndRow As Long)
    m_lEndRow = vlEndRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lEndCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lEndCol() As Long
    lEndCol = m_lEndCol
End Property

Public Property Let lEndCol(ByVal vlEndCol As Long)
    m_lEndCol = vlEndCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lKeyRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lKeyRow() As Long
    lKeyRow = m_lKeyRow
End Property

Public Property Let lKeyRow(ByVal vlKeyRow As Long)
    m_lKeyRow = vlKeyRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lKeyCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lKeyCol() As Long
    lKeyCol = m_lKeyCol
End Property

Public Property Let lKeyCol(ByVal vlKeyCol As Long)
    m_lKeyCol = vlKeyCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetRow() As Long
    lTargetRow = m_lTargetRow
End Property

Public Property Let lTargetRow(ByVal vlTargetRow As Long)
    m_lTargetRow = vlTargetRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetCol() As Long
    lTargetCol = m_lTargetCol
End Property

Public Property Let lTargetCol(ByVal vlTargetCol As Long)
    m_lTargetCol = vlTargetCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lResultRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lResultRow() As Long
    lResultRow = m_lResultRow
End Property

Public Property Let lResultRow(ByVal vlResultRow As Long)
    m_lResultRow = vlResultRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lResultCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lResultCol() As Long
    lResultCol = m_lResultCol
End Property

Public Property Let lResultCol(ByVal vlResultCol As Long)
    m_lResultCol = vlResultCol
End Property


'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sWorkbookName = ""
    m_sWorksheetName = ""
    m_lStartRow = 0
    m_lStartCol = 0
    m_lEndRow = 0
    m_lEndCol = 0
    m_lKeyRow = 0
    m_lKeyCol = 0
    m_lTargetRow = 0
    m_lTargetCol = 0
    m_lResultRow = 0
    m_lResultCol = 0
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sWorksheetName Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : check my value row
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValueRow() As Boolean
    If ws Is Nothing Then
        checkMyValueRow = False
    End If
    If 0 = m_lKeyCol Then
        checkMyValueRow = False
    End If
    checkMyValueRow = True
End Function

'--------------------------------------------------------------------------------
'@brief : check my value col
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValueCol() As Boolean
    If ws Is Nothing Then
        checkMyValueCol = False
    End If
    If 0 = m_lKeyRow Then
        checkMyValueCol = False
    End If
    checkMyValueCol = True
End Function

'--------------------------------------------------------------------------------
'@brief : set sheet info
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub SetSheetInfo()
    If checkMyValue Then
        If sWorkbookName = "" Then
            Set m_wb = ThisWorkbook
            Set m_ws = wb.Worksheets(sWorksheetName)
        ElseIf sWorkbookName <> "" Then
            Set m_wb = Workbooks(sWorkbookName)
            Set m_ws = wb.Worksheets(sWorksheetName)
        Else
            MsgBox "err"
            Exit Sub
        End If
    End If
End Sub

'--------------------------------------------------------------------------------
'@brief : set row and col info
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub SetRowAndColInfo()
    If 0 = m_lEndRow And checkMyValueRow Then
        m_lEndRow = ws.Cells(ws.Rows.Count, m_lKeyCol).End(xlUp).Row
    End If
    If 0 = m_lEndCol And checkMyValueRow Then
        m_lEndCol = ws.Cells(m_lKeyRow, ws.columns.Count).End(xlleft).column
    End If
End Sub
