VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInputFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_lStartRow As Long
Dim m_lInputCol As Long
Dim m_sTargetFolderPath As String

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
Public Property Get lInputCol() As Long
    lInputCol = m_lInputCol
End Property

Public Property Let lInputCol(ByVal vlInputCol As Long)
    m_lInputCol = vlInputCol
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get sTargetFolderPath() As String
    sTargetFolderPath = m_sTargetFolderPath
End Property

Public Property Let sTargetFolderPath(ByVal vsTargetFolderPath As String)
    m_sTargetFolderPath = vsTargetFolderPath
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sWorkbookName = ""
    m_sWorksheetName = ""
    m_lStartRow = 0
    m_sTargetFolderPath = ""
End Sub

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Sub InputFolderName()
    Dim wsTarget As Worksheet
    Dim sDirBuf As String
    Dim lCurRow As Long
    
    If sWorkbookName = "" Then
        Set wsTarget = ThisWorkbook.Worksheets(sWorksheetName)
    ElseIf wsTarget <> "" Then
        Set wsTarget = Workbooks(sWorkbookName).Worksheets(sWorksheetName)
    Else
        MsgBox "err"
        Exit Sub
    End If
    
    sDirBuf = Dir(sTargetFolderPath & "\" & "*.*")
    lCurRow = lStartRow
    Do While sDirBuf <> ""
        wsTarget.Cells(lCurRow, lInputCol).Value = sDirBuf
        sDirBuf = Dir()
        lCurRow = lCurRow + 1
    Loop
    
End Sub
