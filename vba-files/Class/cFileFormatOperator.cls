VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileFormatOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is to operate file format.
' The correspondence between extensions and enum is shown below:
' xlWorkbookNormal              : *.xls
' xlCSV                         : *.csv
' xlXMLSpreadsheet              : *.xml
' xlOpenXMLWorkbook             : *.xlsx
' xlOpenXMLWorkbookMacroEnabled : *.xlsm

Private Enum M_E_EXTENSION
    FILE_XLS
    FILE_CSV
    FILE_XML
    FILE_XLSX
    FILE_XLSM
    FILE_NONE
    FILE_EXTENSION_MAX
End Enum

Private Const M_S_EXTENSION_XLS As String = ".xls"
Private Const M_S_EXTENSION_CSV As String = ".csv"
Private Const M_S_EXTENSION_XML As String = ".xml"
Private Const M_S_EXTENSION_XLSX As String = ".xlsx"
Private Const M_S_EXTENSION_XLSM As String = ".xlsm"
Private Const M_S_EXTENSION_NONE As String = ""

Private m_lExtension() As Long
Private m_sExtension() As String

' Initialize member array.
Private Sub Class_Initialize()
    ReDim m_lExtension(M_E_EXTENSION.FILE_EXTENSION_MAX - 1)
    ReDim m_sExtension(M_E_EXTENSION.FILE_EXTENSION_MAX - 1)
    
    m_lExtension(M_E_EXTENSION.FILE_XLS) = xlWorkbookNormal
    m_lExtension(M_E_EXTENSION.FILE_CSV) = xlCSV
    m_lExtension(M_E_EXTENSION.FILE_XML) = xlXMLSpreadsheet
    m_lExtension(M_E_EXTENSION.FILE_XLSX) = xlOpenXMLWorkbook
    m_lExtension(M_E_EXTENSION.FILE_XLSM) = xlOpenXMLWorkbookMacroEnabled
    m_lExtension(M_E_EXTENSION.FILE_NONE) = xlNone
    
    m_sExtension(M_E_EXTENSION.FILE_XLS) = M_S_EXTENSION_XLS
    m_sExtension(M_E_EXTENSION.FILE_CSV) = M_S_EXTENSION_CSV
    m_sExtension(M_E_EXTENSION.FILE_XML) = M_S_EXTENSION_XML
    m_sExtension(M_E_EXTENSION.FILE_XLSX) = M_S_EXTENSION_XLSX
    m_sExtension(M_E_EXTENSION.FILE_XLSM) = M_S_EXTENSION_XLSM
    m_sExtension(M_E_EXTENSION.FILE_NONE) = M_S_EXTENSION_NONE
End Sub

' Receive enum of file format then return correspond extension of string.
' If param extension is not determine, return xlNone.
Public Function GetFileFormatString(ByVal vvFileFormat As Variant) As String
    Dim bTarget As Boolean: bTarget = False
    Dim lCntArray As Long
    GetFileFormatString = m_sExtension(M_E_EXTENSION.FILE_NONE)
    For lCntArray = LBound(m_lExtension) To UBound(m_lExtension) - 1
        If vbLong = VarType(vvFileFormat) And vvFileFormat = m_lExtension(lCntArray) Then
            bTarget = True
        End If
        If vbString = VarType(vvFileFormat) And lCntArray <> M_E_EXTENSION.FILE_NONE _
        And vvFileFormat Like "*" & m_sExtension(lCntArray) Then
            bTarget = True
        End If
        If bTarget Then
            GetFileFormatString = m_sExtension(lCntArray)
            Exit Function
        End If
    Next lCntArray
End Function

' Receive enum of file format then return correspond extension of enum.
' If param extension is not determine, return "".
Public Function GetFileFormatEnum(ByVal vsFileFormat As String) As Long
    Dim lCntArray As Long
    GetFileFormatEnum = m_lExtension(M_E_EXTENSION.FILE_NONE)
    For lCntArray = LBound(m_sExtension) To UBound(m_sExtension) - 1
        If lCntArray <> M_E_EXTENSION.FILE_NONE _
        And vsFileFormat Like "*" & m_sExtension(lCntArray) Then
            GetFileFormatEnum = m_lExtension(lCntArray)
            Exit Function
        End If
    Next lCntArray
End Function

' Receive file format then if the file format correspond extension return True.
' If param extension is not determine, return False.
Public Function HasExtension(ByVal vsFileName As String) As Boolean
    Dim lCntArray As Long
    For lCntArray = LBound(m_sExtension) To UBound(m_sExtension) - 1
        If lCntArray <> M_E_EXTENSION.FILE_NONE _
        And vsFileName Like "*" & m_sExtension(lCntArray) Then
            HasExtension = True
            Exit Function
        End If
    Next lCntArray
    HasExtension = False
End Function

' Receive file format then remove extension from the param and return.
' If param do not has extension, return original param.
Public Function RemoveExtension(ByVal vlFileName As String) As String
    If Not HasExtension(vlFileName) Then
        RemoveExtension = vlFileName
        Exit Function
    End If
    
    Dim lCntArray As Long
    For lCntArray = LBound(m_sExtension) To UBound(m_sExtension) - 1
        If vlFileName Like "*" & m_sExtension(lCntArray) Then
            RemoveExtension = Replace(vlFileName, m_sExtension(lCntArray), "")
            Exit Function
        End If
    Next lCntArray
    RemoveExtension = vlFileName
End Function
