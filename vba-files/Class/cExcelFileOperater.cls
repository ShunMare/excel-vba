VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExcelFileOperater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''xlWorkbookNormal              ：Microsoft Excelワークブック：*.xls
''xlCSV                         ：CSV(カンマ区切り)：*.csv
''xlXMLSpreadsheet              ：XML スプレッドシート：*.xml
''xlOpenXMLWorkbook             ：XML ブックを開く：*.xlsx
''xlOpenXMLWorkbookMacroEnabled ：マクロを有効にした XML ブックを開く：*.xlsm
Const M_SHEET_NAME_DEFAULT As String = "default"
Dim m_lCreateFileFormat As Long
Dim m_lSourceFileFormat As Long
Dim m_lDestinationFileFormat As Long
Dim m_sCopySheetName As String
Dim m_sChangeSheetName  As String
Dim m_sCreateFolderPath As String
Dim m_sCreateFileName As String
Dim m_sDestinationFolderPath As String
Dim m_sDestinationFileName As String
Dim m_sSourceFolderPath As String
Dim m_sSourceFileName As String

'--------------------------------------------------------------------------------
'@brief : getter letter lCreateFileFormat
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lCreateFileFormat() As Long
    lCreateFileFormat = m_lCreateFileFormat
End Property

Public Property Let lCreateFileFormat(ByVal vlCreateFileFormat As Long)
    m_lCreateFileFormat = vlCreateFileFormat
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lSourceFileFormat
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lSourceFileFormat() As Long
    lSourceFileFormat = m_lSourceFileFormat
End Property

Public Property Let lSourceFileFormat(ByVal vlSourceFileFormat As Long)
    m_lSourceFileFormat = vlSourceFileFormat
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lDestinationFileFormat
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lDestinationFileFormat() As Long
    lDestinationFileFormat = m_lDestinationFileFormat
End Property

Public Property Let lDestinationFileFormat(ByVal vlDestinationFileFormat As Long)
    m_lDestinationFileFormat = vlDestinationFileFormat
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
'@brief : getter letter sChangeSheetName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sChangeSheetName() As String
    sChangeSheetName = m_sChangeSheetName
End Property

Public Property Let sChangeSheetName(ByVal vsChangeSheetName As String)
    m_sChangeSheetName = vsChangeSheetName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sCreateFolderPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sCreateFolderPath() As String
    sCreateFolderPath = m_sCreateFolderPath
End Property

Public Property Let sCreateFolderPath(ByVal vsCreateFolderPath As String)
    m_sCreateFolderPath = vsCreateFolderPath
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
'@brief : getter letter sDestinationFolderPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sDestinationFolderPath() As String
    sDestinationFolderPath = m_sDestinationFolderPath
End Property

Public Property Let sDestinationFolderPath(ByVal vsDestinationFolderPath As String)
    m_sDestinationFolderPath = vsDestinationFolderPath
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sDestinationFileName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sDestinationFileName() As String
    sDestinationFileName = m_sDestinationFileName
End Property

Public Property Let sDestinationFileName(ByVal vsDestinationFileName As String)
    m_sDestinationFileName = vsDestinationFileName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sSourceFolderPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sSourceFolderPath() As String
    sSourceFolderPath = m_sSourceFolderPath
End Property

Public Property Let sSourceFolderPath(ByVal vsSourceFolderPath As String)
    m_sSourceFolderPath = vsSourceFolderPath
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sSourceFileName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sSourceFileName() As String
    sSourceFileName = m_sSourceFileName
End Property

Public Property Let sSourceFileName(ByVal vsSourceFileName As String)
    m_sSourceFileName = vsSourceFileName
End Property

'--------------------------------------------------------------------------------
'@brief : getFileFormat
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Function getFileFormat(ByVal vlFileFormat As Long) As String
    Select Case vlFileFormat
    Case xlWorkbookNormal
        getFileFormat = ".xls"
    Case xlCSV
        getFileFormat = ".csv"
    Case xlXMLSpreadsheet
        getFileFormat = ".xml"
    Case xlOpenXMLWorkbook
        getFileFormat = ".xlsx"
    Case xlOpenXMLWorkbookMacroEnabled
        getFileFormat = ".xlsm"
    Case xlNone
        getFileFormat = ""
    End Select
End Function

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_lCreateFileFormat = 0
    m_sCopySheetName = ""
    m_sChangeSheetName = ""
    m_sCreateFolderPath = ""
    m_sCreateFileName = ""
    m_sDestinationFolderPath = ""
    m_sDestinationFileName = ""
    m_sSourceFolderPath = ""
    m_sSourceFileName = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sCreateFolderPath Then
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
Public Sub CreateFile()
    Dim sCreateFullPath As String
    Dim wbCreate As Workbook
    
    sCreateFullPath = sCreateFolderPath & "\" & sCreateFileName
    If Dir(sCreateFullPath) <> "" Then
        Exit Sub
    End If
    Set wbCreate = Workbooks.Add
    wbCreate.Worksheets(1).Name = M_SHEET_NAME_DEFAULT
    Call wbCreate.SaveAs(Filename:=sCreateFullPath, FileFormat:=lCreateFileFormat)
    Application.DisplayAlerts = False
    wbCreate.Close
    Application.DisplayAlerts = True
End Sub

'--------------------------------------------------------------------------------
'@brief : copy sheet
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub CopySheet()
    Dim sSorceFileFormat As String
    Dim sDestinationFileFormat As String
    Dim sSorceFileFullPath As String
    Dim sDestinationFileFullPath As String
    Dim wbSorce As Workbook
    Dim wbDestination As Workbook
    Dim bIsSorceFile As Boolean
    Dim bIsDestinationFile As Boolean
    
    sSorceFileFormat = getFileFormat(lSourceFileFormat)
    sDestinationFileFormat = getFileFormat(lDestinationFileFormat)
    sSorceFileFullPath = sSourceFolderPath & "\" & sSourceFileName & sSorceFileFormat
    sDestinationFileFullPath = sDestinationFolderPath & "\" & sDestinationFileName & sDestinationFileFormat
    
    bIsSorceFile = False
    bIsDestinationFile = False
    If Not IsWorkbook(sSourceFileName & sSorceFileFormat) Then
        Workbooks.Open sSorceFileFullPath
    End If
    If Not IsWorkbook(sDestinationFileName & sDestinationFileFormat) Then
        Workbooks.Open sDestinationFileFullPath
    End If
    Set wbSorce = Workbooks(sSourceFileName & sSorceFileFormat)
    Set wbDestination = Workbooks(sDestinationFileName & sDestinationFileFormat)
    
    wbSorce.Worksheets(sCopySheetName).Copy before:=wbDestination.Worksheets(1)
    wbDestination.Worksheets(sCopySheetName).Visible = True
    If sChangeSheetName <> "" Then: wbDestination.Worksheets(1).Name = sChangeSheetName
    Application.DisplayAlerts = False
    wbDestination.Worksheets(M_SHEET_NAME_DEFAULT).Delete
    If Not ThisWorkbook.Name = wbSorce.Name Then
        wbSorce.Close
    End If
    If Not ThisWorkbook.Name = wbDestination.Name Then
        wbDestination.Save
        wbDestination.Close
    End If
    Application.DisplayAlerts = True
End Sub

'--------------------------------------------------------------------------------
'@brief : is workbook
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function IsWorkbook(ByVal sWorkbookName As String) As Boolean
    Dim wbTarget As Workbook
    For Each wbTarget In Workbooks
        If wbTarget.Name = sWorkbookName Then
            IsWorkbook = True
            Exit Function
        End If
    Next wbTarget
    IsWorkbook = False
End Function

'--------------------------------------------------------------------------------
'@brief : copy book
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub CopyBook()
    Dim sSorceFileFormat As String
    Dim sDestinationFileFormat As String
    Dim sSorceFileFullPath As String
    Dim sDestinationFileFullPath As String
    Dim wbSorce As Workbook
    Dim bIsSorceFile As Boolean

    sSorceFileFormat = getFileFormat(lSourceFileFormat)
    sDestinationFileFormat = getFileFormat(lDestinationFileFormat)
    sSorceFileFullPath = sSourceFolderPath & "\" & sSourceFileName & sSorceFileFormat
    sDestinationFileFullPath = sDestinationFolderPath & "\" & sDestinationFileName & sDestinationFileFormat
    
    bIsSorceFile = False
    If Not IsWorkbook(sSourceFileName & sSorceFileFormat) Then
        Workbooks.Open sSorceFileFullPath
    End If
    Set wbSorce = Workbooks(sSourceFileName & sSorceFileFormat)
    If Dir(sDestinationFileFullPath) <> "" Then
        Exit Sub
    End If
    wbSorce.SaveCopyAs sDestinationFileFullPath
    Application.DisplayAlerts = False
    If Not ThisWorkbook.Name = wbSorce.Name Then
        wbSorce.Close
    End If
    Application.DisplayAlerts = True
End Sub
