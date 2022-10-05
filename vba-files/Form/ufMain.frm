VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufMain 
   Caption         =   "集計及び日報"
   ClientHeight    =   3101
   ClientLeft      =   98
   ClientTop       =   406
   ClientWidth     =   6811
   OleObjectBlob   =   "ufMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
'@brief : UserForm Initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Set g_siConfig = New cSheetInfo
    With g_siConfig
        .sWorksheetName = G_S_SHEET_NAME_CONFIG
        Call .SetSheetInfo
    End With
    Set g_Config = New cConfig
    Call init
End Sub

'--------------------------------------------------------------------------------
'@brief : init
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub init()
    Dim IsDriveChecker As New cIsDriveChecker
    With IsDriveChecker
        Call .checkIsDrive
        g_siConfig.ws.Range(G_S_CONFIG_LOCAL_DRIVE) = .sDrive
    End With
    With g_Config
        .sDROExportPath = g_siConfig.ws.Range(G_S_CONFIG_DR_ORI_EXPORT_PATH)
        .sLocalDrive = g_siConfig.ws.Range(G_S_CONFIG_LOCAL_DRIVE)
        Me.tbFileFullPath = .sDROExportPath
    End With
    If cbTargetAll Then
        Dim lCnt As Long
        For lCnt = 0 To E_DIVISION.DIV_NUM - 1
            g_bTargetDivision(lCnt) = True
        Next lCnt
    Else
        g_bTargetDivision(E_DIVISION.DIV_1) = Me.cbDivision1
        g_bTargetDivision(E_DIVISION.DIV_2) = Me.cbDivision2
        g_bTargetDivision(E_DIVISION.DIV_3) = Me.cbDivision3
        g_bTargetDivision(E_DIVISION.DIV_4) = Me.cbDivision4
        g_bTargetDivision(E_DIVISION.DIV_ALL) = Me.cbDivisionAll
    End If
End Sub

'--------------------------------------------------------------------------------
'@brief : cbGetFileFullPath click
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub cbGetFileFullPath_Click()
    Dim FolderPathInputter As New cFolderPathInputter
    Dim sFolderPath As String
    With FolderPathInputter
        sFolderPath = .GetFolderPath
        ThisWorkbook.Worksheets(G_S_SHEET_NAME_CONFIG).Range(G_S_CONFIG_DR_ORI_EXPORT_PATH) = sFolderPath
    End With
    Call init
End Sub

'--------------------------------------------------------------------------------
'@brief : UserForm Initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub cbTargetAll_Change()
    If cbTargetAll Then
        cbDivision1.Enabled = False
        cbDivision2.Enabled = False
        cbDivision3.Enabled = False
        cbDivision4.Enabled = False
        cbDivisionAll.Enabled = False
    Else
        cbDivision1.Enabled = True
        cbDivision2.Enabled = True
        cbDivision3.Enabled = True
        cbDivision4.Enabled = True
        cbDivisionAll.Enabled = True
    End If
End Sub

'--------------------------------------------------------------------------------
'@brief : cbStart click
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub cbStart_Click()
    Call init
    If Not needsMirigrant Then Exit Sub
    Call controlForm(False)
    Call SummarizeData
    Call controlForm(True)
    Call fini
End Sub

'--------------------------------------------------------------------------------
'@brief : needsMirigrant
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Function needsMirigrant()
    Dim lCntLoop As Long
    Dim lCnt As Long: lCnt = 0
    Dim MessageShower As New cMessageShower
    
    If g_Config.sDROExportPath = "" Then
        MessageShower.ShowCriticalMsg (G_S_MSG0001)
        needsMirigrant = False
        Exit Function
    End If
    If g_Config.sLocalDrive = "" Then
        MessageShower.ShowCriticalMsg (G_S_MSG0008)
        needsMirigrant = False
        Exit Function
    End If
    If Dir(g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT, vbDirectory) <> "" Then
        MessageShower.sReplaceWord1 = g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT
        MessageShower.ShowCriticalMsg (G_S_MSG0009)
        needsMirigrant = False
        Exit Function
    Else
        MkDir g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT
    End If
    For lCntLoop = 0 To E_DIVISION.DIV_NUM - 1
        If Not g_bTargetDivision(lCntLoop) Then lCnt = lCnt + 1
    Next lCntLoop
    If lCnt = E_DIVISION.DIV_NUM Then
        MessageShower.ShowExclamationMsg (G_S_MSG0007)
        needsMirigrant = False
        Exit Function
    End If
    needsMirigrant = True
End Function

'--------------------------------------------------------------------------------
'@brief : cbFinish click
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub cbFinish_Click()
    Unload ufMain
End Sub

'--------------------------------------------------------------------------------
'@brief : fini
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub fini()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFolder g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT, _
    g_Config.sLocalDrive & "\" & G_S_FOLDER_NAME_DAIRY_REPORT & "_" _
    & Format(Year(Now()), "0000") & Format(Month(Now()), "00") & Format(Day(Now()), "00") _
    & Format(Hour(Now()), "00") & Format(Minute(Now()), "00") & Format(Second(Now()), "00")
End Sub

'--------------------------------------------------------------------------------
'@brief : ControlForm
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub controlForm(ByVal vbFlag As Boolean)
    Dim oElem As Object

    With ufMain
        For Each oElem In .Controls
            oElem.Enabled = vbFlag
        Next
    End With
End Sub
