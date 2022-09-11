VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cReplaceWords"
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
Dim m_sTargetWord As String
Dim m_sReplaceWord As String

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
'@brief : getter letter sTargetWord
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetWord() As String
    sTargetWord = m_sTargetWord
End Property

Public Property Let sTargetWord(ByVal vsTargetWord As String)
    m_sTargetWord = vsTargetWord
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sReplaceWord
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sReplaceWord() As String
    sReplaceWord = m_sReplaceWord
End Property

Public Property Let sReplaceWord(ByVal vsReplaceWord As String)
    m_sReplaceWord = vsReplaceWord
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
    m_lEndCol = 0
    m_lKeyCol = 0
    m_sTargetWord = ""
    m_sReplaceWord = ""
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
    If "" = m_lStartRow Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_lStartCol Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_lEndCol Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_lKeyCol Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_sTargetWord Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_sReplaceWord Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : replace words
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function ClearRangeContents()
    Dim wsTarget As Worksheet
    Dim rTarget As Range
    Dim lEndRow As Long
    
    checkMyValue
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
        rTarget.Replace What:=sTargetWord, Replacement:=sReplaceWord, LookAt:=xlWhole, MatchCase:=True
    End With
End Function
