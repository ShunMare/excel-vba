VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileCreater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const M_SHEET_NAME_DEFAULT As String = "default"
Dim m_sTargetFolderPath As String
Dim m_sCopySheetName As String
Dim m_sCreateFileName As String
Dim m_sTargetExtension  As String
Dim m_lTargetExtension  As Long
Dim m_sChangeFileName  As String
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
'@brief : getter letter sCopySheetName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sCopySheetName() As String
    sCopySheetName = m_sCopySheetName
End Property

Public Property Let sCopySheetName(ByVal vsCopySheetName As String)
    m_sCopySheetName = vsCopySheetName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sCreateFileName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sCreateFileName() As String
    sCreateFileName = m_sCreateFileName
End Property

Public Property Let sCreateFileName(ByVal vsCreateFileName As String)
    m_sCreateFileName = vsCreateFileName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetExtension
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetExtension() As Long
    lTargetExtension = m_lTargetExtension
End Property

Public Property Let lTargetExtension(ByVal vlTargetExtension As Long)
    m_lTargetExtension = vlTargetExtension
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sTargetExtension
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetExtension() As String
    sTargetExtension = m_sTargetExtension
End Property

Public Property Let sTargetExtension(ByVal vsTargetExtension As String)
    m_sTargetExtension = vsTargetExtension
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sChangeFileName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sChangeFileName() As String
    sChangeFileName = m_sChangeFileName
End Property

Public Property Let sChangeFileName(ByVal vsChangeFileName As String)
    m_sChangeFileName = vsChangeFileName
End Property
'--------------------------------------------------------------------------------
'@brief : get extension
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub getExtention()
    Select Case lTargetExtension
    Case E_EXTENSION.XLS
        sTargetExtension = ".xls"
    Case E_EXTENSION.XLSX
        sTargetExtension = ".xlsx"
    Case E_EXTENSION.XLSM
        sTargetExtension = ".xlsm"
    Case E_EXTENSION.TXT
        sTargetExtension = ".txt"
    Case E_EXTENSION.PDF
        sTargetExtension = ".pdf"
    End Select
End Sub
'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sTargetFolderPath = ""
    m_sCopySheetName = ""
    m_sChangeFileName = ""
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
'@brief : create excel file
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub createExcelFile()
    Dim sCreateFullPath As String
    Dim wbCreate As Workbook
    
    Call getExtention
    sCreateFullPath = sTargetFolderPath & "\" & sCreateFileName & sTargetExtension
    If Dir(sCreateFullPath) <> "" Then
        Exit Sub
    End If
    Set wbCreate = Workbooks.Add
    wbCreate.Worksheets(1).Name = M_SHEET_NAME_DEFAULT
    wbCreate.SaveAs sCreateFullPath
End Sub

'--------------------------------------------------------------------------------
'@brief : copy my sheet
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub copyMySheet()
    Dim wbOpen As Workbook
    
    Set wbOpen = Workbooks(sCreateFileName & sTargetExtension)
    ThisWorkbook.Worksheets(sCopySheetName).Copy before:=wbOpen.Worksheets(1)
    If sChangeFileName <> "" Then: wbOpen.Worksheets(1).Name = sChangeFileName
    Application.DisplayAlerts = False
    wbOpen.Worksheets(M_SHEET_NAME_DEFAULT).Delete
    Application.DisplayAlerts = True
    wbOpen.Save
End Sub

'--------------------------------------------------------------------------------
'@brief : close created file
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub closeCreatedFile()
    Dim wbOpen As Workbook
    
    Set wbOpen = Workbooks(sCreateFileName & sTargetExtension)
    wbOpen.Close
End Sub
