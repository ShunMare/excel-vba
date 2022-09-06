VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInputResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_lStartRow As Long
Dim m_lTargetCol As Long

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
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sWorkbookName = ""
    m_sWorksheetName = ""
    m_lStartRow = 0
    m_lTargetCol = 0
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sWorkbookName Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_sWorksheetName Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lStartRow Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lTargetCol Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : input result
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function InputResult(ByVal vlTargetRow As Long, ByVal vsMsg As String)
    Dim rTarget As Range
    
    If sWorkbookName = "" Then
        Set rTarget = ThisWorkbook.Worksheets(sWorksheetName).Cells(vlTargetRow, lTargetCol)
    ElseIf sWorkbookName <> "" Then
        Set rTarget = Workbooks(sWorkbookName).Worksheets(sWorksheetName).Cells(vlTargetRow, lTargetCol)
    Else
        MsgBox "err"
        Exit Function
    End If
    
    If vsMsg <> "" Then
        rTarget = vsMsg
    Else
        rTarget = "TRUE"
    End If
End Function
