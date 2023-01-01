VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsCngFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Private Const M_L_START_ROW As Long = 2
' Private Const M_L_CURRENT_FILE_NAME_COL As Long = 1
' Private Const M_L_CNG_FILE_NAME_COL As Long = 2
' Private Const M_L_RESULT_COL As Long = 3

' Private m_siChgFileName As cSheetInfo

' '--------------------------------------------------------------------------------
' '@brief : init
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Private Sub init()
'     Set m_siChgFileName = New cSheetInfo
'     With m_siChgFileName
'         '.sWorksheetName = G_S_CNG_FILE_NAME
'         Call .SetSheetInfo
'         .lStartRow = M_L_START_ROW
'         .lStartCol = M_L_CURRENT_FILE_NAME_COL
'         .lEndCol = M_L_RESULT_COL
'         .lKeyCol = M_L_CURRENT_FILE_NAME_COL
'         Call .SetRowAndColInfo
'     End With
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnInputFolderPath
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnInputFolderPath_Click()
'     Call init
    
'     Dim FolderPathInputter As New cFolderPathInputter
'     With FolderPathInputter
'         '.sCellName = G_S_TARGET_FOLDER_PATTH
'         Call .InputFolderPath(m_siChgFileName)
'     End With
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnInputFileName
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnInputFileName_Click()
'     Call init
    
'     With m_siChgFileName
'         .lTargetCol = M_L_CURRENT_FILE_NAME_COL
'         .SetRowAndColInfo
'     End With
    
'     Dim FileNameInputter As New cFileNameInputter
'     With FileNameInputter
'         '.sTargetFolderPath = Range(G_S_TARGET_FOLDER_PATTH)
'         Call .InputFolderName(m_siChgFileName)
'     End With
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnCngFileName
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnCngFileName_Click()
'     Call init
    
'     Dim FileNameChanger As New cFileNameChanger
'     With FileNameChanger
'         .lCurFileNameCol = M_L_CURRENT_FILE_NAME_COL
'         .lChangeFileNameCol = M_L_CNG_FILE_NAME_COL
'         '.sTargetFolderPath = Range(G_S_TARGET_FOLDER_PATTH)
'         Call .CngFileName(m_siChgFileName)
'     End With
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnClear
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnClear_Click()
'     Call init

'     Dim RangeClearer As New cRangeClearer
'     With RangeClearer
'         Call .ClearRangeContents(m_siChgFileName)
'     End With
' End Sub
