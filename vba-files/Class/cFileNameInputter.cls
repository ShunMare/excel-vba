VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileNameInputter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sTargetFolderPath As String

'--------------------------------------------------------------------------------
'@brief : getter letter sTargetFolderPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetFolderPath() As String
    sTargetFolderPath = m_sTargetFolderPath
End Property

Public Property Let sTargetFolderPath(ByVal vsTargetFolderPath As String)
    m_sTargetFolderPath = vsTargetFolderPath
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sTargetFolderPath = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sTargetFolderPath Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : input folder name
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub InputFolderName(ByVal siTarget As cSheetInfo)
    Dim sDirBuf As String
    Dim lCurRow As Long

    If Not checkMyValue Then
        Exit Sub
    End If

    sDirBuf = Dir(sTargetFolderPath & "\" & "*.*")
    lCurRow = siTarget.lStartRow
    Do While sDirBuf <> ""
        siTarget.ws.Cells(lCurRow, siTarget.lTargetCol).Value = sDirBuf
        sDirBuf = Dir()
        lCurRow = lCurRow + 1
    Loop
End Sub
