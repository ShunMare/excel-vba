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

Private Const M_L_START_ROW = 2
Private Const M_L_CURRENT_FILE_NAME_COL = 1
Private Const M_L_CNG_FILE_NAME_COL = 2
Private Const M_L_RESULT_COL = 3

'--------------------------------------------------------------------------------
'@brief : click btnInputFolderPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnInputFolderPath_Click()
    Dim InputFolderPath As cInputFolderPath
    Set InputFolderPath = New cInputFolderPath
    
    With InputFolderPath
        .sWorksheetName = G_S_CNG_FILE_NAME
        .sCellName = G_S_TARGET_FOLDER_PATTH
        Call .InputFolderPath
    End With
End Sub

'--------------------------------------------------------------------------------
'@brief : click btnInputFileName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnInputFileName_Click()
    Dim InputFileName As New cInputFileName
    Set InputFileName = New cInputFileName
    
    With InputFileName
        .sWorksheetName = Me.Name
        .lStartRow = M_L_START_ROW
        .lInputCol = M_L_CURRENT_FILE_NAME_COL
        .sTargetFolderPath = Range(G_S_TARGET_FOLDER_PATTH)
        Call .InputFolderName
    End With
End Sub

'--------------------------------------------------------------------------------
'@brief : click btnCngFileName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnCngFileName_Click()
    Dim CngFileName As New cCngFileName
    Set CngFileName = New cCngFileName
    
    With CngFileName
        .sWorksheetName = Me.Name
        .lStartRow = M_L_START_ROW
        .lCurFileNameCol = M_L_CURRENT_FILE_NAME_COL
        .lChangeFileNameCol = M_L_CNG_FILE_NAME_COL
        .lResultCol = M_L_RESULT_COL
        .sTargetFolderPath = Range(G_S_TARGET_FOLDER_PATTH)
        Call .CngFileName
    End With
End Sub

'--------------------------------------------------------------------------------
'@brief : click btnClear
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnClear_Click()
    Dim ClearContents As New cClearContents
    Set ClearContents = New cClearContents
    
    With ClearContents
        .sWorksheetName = Me.Name
        .lStartRow = M_L_START_ROW
        .lStartCol = M_L_CURRENT_FILE_NAME_COL
        .lEndCol = M_L_RESULT_COL
        .lKeyCol = M_L_CURRENT_FILE_NAME_COL
        Call .ClearRangeContents
    End With
End Sub
