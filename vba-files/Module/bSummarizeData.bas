Attribute VB_Name = "bSummarizeData"
Option Explicit

Private m_sDivision(E_DIVISION.DIV_NUM) As String
Private m_sLargeItems(E_LARGE_ITEM.LARGE_ITEM_NUM) As String
Private m_lStartRowOfRange(E_LARGE_ITEM.LARGE_ITEM_NUM) As Long
Private m_siDailyReportTmp As cSheetInfo
Private m_siExport As cSheetInfo
Private m_siClearRange As cSheetInfo
Private m_riInport As cRangeInfo_DR
Private m_riExport As cRangeInfo_DR
Private m_MessageShower As cMessageShower
Private m_sMsg As String
Private Const m_lEndRowCorrection As Long = 2

'--------------------------------------------------------------------------------
'@brief : summarize data
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub SummarizeData()
    On Error GoTo LabelError
    m_sMsg = ""
    Application.ScreenUpdating = False
    Call init
    Call needsMigrate
    Call setRangeInfo(m_siDailyReportTmp, m_riInport)
    Call clearRange
    
    Dim lCntLoop As Long
    For lCntLoop = 0 To UBound(g_bTargetDivision)
        If g_bTargetDivision(lCntLoop) Then
            Call operateExportFile(lCntLoop)
            Call CreateFile(lCntLoop)
            Call clearRange
        End If
    Next lCntLoop
    
    Dim sTmpMsg As String: sTmpMsg = ""
    For lCntLoop = 0 To UBound(g_bTargetDivision)
        If g_bTargetDivision(lCntLoop) Then
            sTmpMsg = sTmpMsg & m_sDivision(lCntLoop)
        End If
    Next lCntLoop
    If sTmpMsg = "" Then
        m_MessageShower.ShowExclamationMsg (G_S_MSG0006)
    Else
        m_MessageShower.sReplaceWord1 = sTmpMsg
        m_MessageShower.ShowInformationMsg (G_S_MSG0005)
    End If
    Application.ScreenUpdating = True
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------
'@brief : init
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub init()
    On Error GoTo LabelError
    m_sMsg = ""
    Set m_MessageShower = New cMessageShower
    Set m_riInport = New cRangeInfo_DR
    Set m_riExport = New cRangeInfo_DR

    Set m_siDailyReportTmp = New cSheetInfo
    With m_siDailyReportTmp
        .sWorksheetName = G_S_SHEET_NAME_DAIRY_REPORT_ORI
        Call .SetSheetInfo
        .lStartRow = 1
        .lStartCol = 1
        .lKeyRow = 4: .lKeyCol = 1
        .lTargetCol = 3
        Call .SetRowAndColInfo
        .lEndRow = .lEndRow + m_lEndRowCorrection
    End With
    
    Set m_siClearRange = New cSheetInfo
    
    m_sDivision(E_DIVISION.DIV_1) = G_S_DIVISION_1
    m_sDivision(E_DIVISION.DIV_2) = G_S_DIVISION_2
    m_sDivision(E_DIVISION.DIV_3) = G_S_DIVISION_3
    m_sDivision(E_DIVISION.DIV_4) = G_S_DIVISION_4
    m_sDivision(E_DIVISION.DIV_ALL) = G_S_DIVISION_ALL
    
    m_sLargeItems(E_LARGE_ITEM.DAIRY_REPORT) = G_S_LARGE_TITLE_DAIRY_REPORT
    m_sLargeItems(E_LARGE_ITEM.POTENTIAL_COSTOMER_KPI) = G_S_LARGE_TITLE_POTENTIAL_COSTOMER_KPI
    m_sLargeItems(E_LARGE_ITEM.NEW_KPI) = G_S_LARGE_TITLE_NEW_KPI
    m_sLargeItems(E_LARGE_ITEM.TIME_KPI) = G_S_LARGE_TITLE_TIME_KPI
    m_sLargeItems(E_LARGE_ITEM.GOAL) = G_S_LARGE_TITLE_GOAL
    
    m_lStartRowOfRange(E_LARGE_ITEM.DAIRY_REPORT) = 0
    m_lStartRowOfRange(E_LARGE_ITEM.POTENTIAL_COSTOMER_KPI) = 3
    m_lStartRowOfRange(E_LARGE_ITEM.NEW_KPI) = 2
    m_lStartRowOfRange(E_LARGE_ITEM.TIME_KPI) = 3
    m_lStartRowOfRange(E_LARGE_ITEM.GOAL) = 0
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub

'--------------------------------------------------------------------------------
'@brief : needs Migrate
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub needsMigrate()
    On Error GoTo LabelError
    Dim lCntLoop As Long
    Dim sTargetFullPath As String
    Dim sTargetFileName As String
        
    For lCntLoop = LBound(g_bTargetDivision) To UBound(g_bTargetDivision)
        If g_bTargetDivision(lCntLoop) Then
            sTargetFileName = "*" & m_sDivision(lCntLoop) & "*" & G_S_SUMMARIZING & "*"
            If Dir(g_Config.sDROExportPath & "\" & sTargetFileName) <> "" Then
                m_MessageShower.sReplaceWord1 = m_sDivision(lCntLoop) & G_S_SUMMARIZING
                m_MessageShower.ShowInformationMsg (G_S_MSG0003)
                m_MessageShower.CleanUpWords
                g_bTargetDivision(lCntLoop) = False
            End If
        End If
    Next lCntLoop
    
    Dim sFileNameArray() As String
    For lCntLoop = LBound(g_bTargetDivision) To UBound(g_bTargetDivision)
        If g_bTargetDivision(lCntLoop) And lCntLoop <> E_DIVISION.DIV_ALL Then
            Dim FileNameGetter As New cFileNameInputter
            With FileNameGetter
                .sTargetFolderPath = Trim(g_Config.sDROExportPath)
                .lTargetExtension = E_EXTENSION.XLS
                .sFilterWord = m_sDivision(lCntLoop)
                .sNotTargetWord = G_S_SUMMARIZING
                sFileNameArray = .GetFileNameArray
            End With
            Set FileNameGetter = Nothing
            If sFileNameArray(0) = G_S_NOT_EXIST Then
                m_MessageShower.sReplaceWord1 = m_sDivision(lCntLoop)
                m_MessageShower.sReplaceWord2 = m_sDivision(lCntLoop) & G_S_SUMMARIZING
                m_MessageShower.ShowInformationMsg (G_S_MSG0004)
                m_MessageShower.CleanUpWords
                g_bTargetDivision(lCntLoop) = False
            End If
        End If
    Next lCntLoop
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub

'--------------------------------------------------------------------------------
'@brief : set range info
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub setRangeInfo(ByVal siTarget As cSheetInfo, ByRef riTarget As cRangeInfo_DR)
    On Error GoTo LabelError
    Dim lCurRow As Long
    Dim lCntArray As Long
    lCntArray = 0
    
    Set riTarget = New cRangeInfo_DR
    Dim bFirst As Boolean
    Dim lRangeStartRow As Long
    Dim lRangeStartCol As Long
    Dim lRangeEndRow As Long
    Dim lRangeEndCol As Long
    Dim rTmp As Range
    bFirst = True
    lRangeStartCol = siTarget.lStartCol
    lRangeEndCol = siTarget.lEndCol
    For lCurRow = siTarget.lStartRow To siTarget.lEndRow
        If siTarget.ws.Cells(lCurRow, siTarget.lKeyCol) = m_sLargeItems(lCntArray) Then
            If bFirst Then
                bFirst = False
            Else
                lRangeEndRow = IIf(lCntArray <> E_LARGE_ITEM.LARGE_ITEM_NUM, lCurRow - 1, siTarget.lEndRow)
                Set rTmp = siTarget.ws.Range(siTarget.ws.Cells(lRangeStartRow, lRangeStartCol), _
                siTarget.ws.Cells(lRangeEndRow, lRangeEndCol))
                Call riTarget.SetNumberRange(lCntArray - 1, rTmp)
            End If
        ''judge should exit (or continue)
            If lRangeEndRow = siTarget.lEndRow Then
                Exit Sub
            Else
                lRangeStartRow = lCurRow
                lCntArray = lCntArray + 1
            End If
        End If
    Next lCurRow
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub

'--------------------------------------------------------------------------------
'@brief : clear range
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub clearRange()
    On Error GoTo LabelError
    Dim lStartRow As Long
    Dim lStartCol As Long
    Dim lEndCol As Long
    Dim lEndRow As Long
    lStartCol = 3
    
    Dim lCntArray As Long
    For lCntArray = 0 To E_LARGE_ITEM.LARGE_ITEM_NUM - 1
        If lCntArray = E_LARGE_ITEM.DAIRY_REPORT Or lCntArray = E_LARGE_ITEM.GOAL Then
            GoTo Continue
        End If
        With m_riInport.ReturnNumberRange(lCntArray)
            lStartRow = m_lStartRowOfRange(lCntArray): lEndRow = .Rows.Count
            lEndCol = .Columns.Count
            .Parent.Range(.Cells(lStartRow, lStartCol), .Cells(lEndRow, lEndCol)).ClearContents
        End With
Continue:
    Next lCntArray
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub

'--------------------------------------------------------------------------------
'@brief : operate export file
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub operateExportFile(ByVal lTargetDivision As Long)
    On Error GoTo LabelError
    Dim sFileNameArray() As String
    Dim FileNameGetter As New cFileNameInputter
    With FileNameGetter
        .sTargetFolderPath = Trim(g_Config.sDROExportPath)
        .lTargetExtension = E_EXTENSION.XLS
        .sFilterWord = IIf(lTargetDivision = E_DIVISION.DIV_ALL, "", m_sDivision(lTargetDivision))
        .sNotTargetWord = G_S_SUMMARIZING
        sFileNameArray = .GetFileNameArray
    End With
    
    Dim lCntFileNum As Long
    For lCntFileNum = LBound(sFileNameArray) To UBound(sFileNameArray)
        Workbooks.Open FileNameGetter.sTargetFolderPath & "\" & sFileNameArray(lCntFileNum)
        Set m_siExport = New cSheetInfo
        With m_siExport
            .sWorkbookName = sFileNameArray(lCntFileNum)
            .sWorksheetName = G_S_SHEET_NAME_DAIRY_REPORT
            Call .SetSheetInfo
            .lStartRow = 1
            .lStartCol = 1
            .lKeyRow = 4: .lKeyCol = 1
            .lTargetCol = 3
            Call .SetRowAndColInfo
            .lEndRow = .lEndRow + m_lEndRowCorrection
            Application.Windows(.wb.Name).Visible = False
        End With
        Call setRangeInfo(m_siExport, m_riExport)
    ''inport value
        Call inportValue
        Set m_siExport = Nothing
        Workbooks(sFileNameArray(lCntFileNum)).Close SaveChanges:=False
    Next lCntFileNum
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub

'--------------------------------------------------------------------------------
'@brief : inport value
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub inportValue()
    On Error GoTo LabelError
    Dim InportValueGetter As New cValueInRangeGetter
    Dim ExportValueGetter As New cValueInRangeGetter
    With InportValueGetter
        .lItemStartCol = 1: .lItemEndCol = 2
        .lTargetCol = 3
    End With
    With ExportValueGetter
        .lItemStartCol = 1: .lItemEndCol = 2
        .lTargetCol = 3
    End With
    
    Dim lCntArray As Long
    For lCntArray = 0 To E_LARGE_ITEM.LARGE_ITEM_NUM - 1
        If lCntArray = E_LARGE_ITEM.DAIRY_REPORT Or lCntArray = E_LARGE_ITEM.GOAL Then
            GoTo Continue
        End If
        InportValueGetter.lItemStartRow = m_lStartRowOfRange(lCntArray)
        ExportValueGetter.lItemStartRow = m_lStartRowOfRange(lCntArray)
        
        Set InportValueGetter.rTarget = m_riInport.ReturnNumberRange(lCntArray)
        Call InportValueGetter.SetMyValue
        Set ExportValueGetter.rTarget = m_riExport.ReturnNumberRange(lCntArray)
        Call ExportValueGetter.SetMyValue
        
        Dim lStartRow As Long
        Dim lEndRow As Long
        Dim lCurRow As Long
        lStartRow = InportValueGetter.lItemStartRow
        lEndRow = InportValueGetter.lItemEndRow
        For lCurRow = lStartRow To lEndRow
            InportValueGetter.lTargetRow = lCurRow
            ExportValueGetter.sTargetItem = InportValueGetter.GetItemInRange
            InportValueGetter.rTarget.Cells(lCurRow, InportValueGetter.lTargetCol) = _
            InportValueGetter.rTarget.Cells(lCurRow, InportValueGetter.lTargetCol) + ExportValueGetter.GetValueInRange
        Next lCurRow
Continue:
    Next lCntArray
    Exit Sub

LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub

'--------------------------------------------------------------------------------
'@brief : create file
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub CreateFile(ByVal lTargetDivision As Long)
    On Error GoTo LabelError
    
    Dim sCreateFileName As String
    sCreateFileName = G_S_INSIDE_DAIRY_REPORT & "-" & m_sDivision(lTargetDivision) & "-" & G_S_SUMMARIZING
    Dim ExcelFileOperater As New cExcelFileOperater
    With ExcelFileOperater
        .sCreateFolderPath = g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT
        .sCreateFileName = sCreateFileName
        .lCreateFileFormat = xlOpenXMLWorkbook
        Call .CreateFile
        .sCopySheetName = G_S_SHEET_NAME_DAIRY_REPORT_ORI
        .sChangeSheetName = G_S_SHEET_NAME_DAIRY_REPORT
        .sSourceFolderPath = ThisWorkbook.Path
        .sSourceFileName = ThisWorkbook.Name
        .lSourceFileFormat = xlNone
        .sDestinationFolderPath = g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT
        .sDestinationFileName = sCreateFileName
        .lDestinationFileFormat = xlOpenXMLWorkbook
        Call .CopySheet
        .sSourceFolderPath = g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT
        .sSourceFileName = sCreateFileName
        .lSourceFileFormat = xlOpenXMLWorkbook
        .sDestinationFolderPath = g_Config.sDROExportPath
        .sDestinationFileName = sCreateFileName
        .lDestinationFileFormat = xlOpenXMLWorkbook
        Call .CopyBook
    End With
    Exit Sub
    
LabelError:
    If m_sMsg = "" Then
        m_MessageShower.ShowCriticalMsg (Err.Description)
        Err.Clear
    Else
        m_MessageShower.ShowCriticalMsg (m_sMsg)
    End If
    Application.ScreenUpdating = True
    End
End Sub



