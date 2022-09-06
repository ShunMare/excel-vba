VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCngFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_sTargetFolderPath As String
Dim m_lStartRow As Long
Dim m_lChangeFileNameCol As Long
Dim m_lCurFileNameCol As Long
Dim m_lResultCol As Long

'--------------------------------------------------------------------------------
'@brief : getter letter sWorkbookName 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sWorkbookName () As String
    sWorkbookName  = m_sWorkbookName 
End Property

Public Property Let sWorkbookName (ByVal vsWorkbookName  As String)
    m_sWorkbookName  = vsWorkbookName 
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sWorksheetName 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sWorksheetName () As String
    sWorksheetName  = m_sWorksheetName 
End Property

Public Property Let sWorksheetName (ByVal vsWorksheetName  As String)
    m_sWorksheetName  = vsWorksheetName 
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sTargetFolderPath 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetFolderPath () As String
    sTargetFolderPath  = m_sTargetFolderPath 
End Property

Public Property Let sTargetFolderPath (ByVal vsTargetFolderPath  As String)
    m_sTargetFolderPath  = vsTargetFolderPath 
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lStartRow 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lStartRow () As Long
    lStartRow  = m_lStartRow 
End Property

Public Property Let lStartRow (ByVal vlStartRow  As Long)
    m_lStartRow  = vlStartRow 
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lChangeFileNameCol 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lChangeFileNameCol () As Long
    lChangeFileNameCol  = m_lChangeFileNameCol 
End Property

Public Property Let lChangeFileNameCol (ByVal vlChangeFileNameCol  As Long)
    m_lChangeFileNameCol  = vlChangeFileNameCol 
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lCurFileNameCol 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lCurFileNameCol () As Long
    lCurFileNameCol  = m_lCurFileNameCol 
End Property

Public Property Let lCurFileNameCol (ByVal vlCurFileNameCol  As Long)
    m_lCurFileNameCol  = vlCurFileNameCol 
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lResultCol 
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lResultCol () As Long
    lResultCol  = m_lResultCol 
End Property

Public Property Let lResultCol (ByVal vlResultCol  As Long)
    m_lResultCol  = vlResultCol 
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sWorkbookName  = ""
    m_sWorksheetName  = ""
    m_sTargetFolderPath  = ""
    m_lStartRow  = 0
    m_lChangeFileNameCol  = 0
    m_lCurFileNameCol  = 0
    m_lResultCol  = 0
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sWorkbookName  Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_sWorksheetName  Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_sTargetFolderPath  Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lStartRow  Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lChangeFileNameCol  Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lCurFileNameCol  Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lResultCol  Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : change file name
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub CngFileName()
    Dim wsTarget As Worksheet
    Dim lCurRow As Long
    Dim lEndRow As Long
    Dim sTargetFilePath As String
    Dim sTargetFileName As String
    Dim sChangeFileName As String
    Dim oFileSystemObject As Object
    Dim InputResult As cInputResult
    
    Set InputResult = New cInputResult
    With InputResult
        .sWorksheetName = sWorksheetName
        .lTargetCol = lResultCol
    End With
    
    If sWorkbookName = "" Then
        Set wsTarget = ThisWorkbook.Worksheets(sWorksheetName)
    ElseIf wsTarget <> "" Then
        Set wsTarget = Workbooks(sWorkbookName).Worksheets(sWorksheetName)
    Else
        MsgBox "err"
        Exit Sub
    End If
    
    lEndRow = wsTarget.Cells(wsTarget.Rows.Count, lCurFileNameCol).End(xlUp).Row
    Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    If m_lResultCol = 0 Then
    End If
    
    For lCurRow = lStartRow To lEndRow
        sTargetFileName = wsTarget.Cells(lCurRow, lCurFileNameCol)
        sTargetFilePath = sTargetFolderPath & "\" & sTargetFileName
        sChangeFileName = wsTarget.Cells(lCurRow, lChangeFileNameCol)
        
        If Not sChangeFileName Like "*.*" Then
            Call InputResult.InputResult(lCurRow, "not extension")
        ElseIf Not oFileSystemObject.FileExists(sTargetFilePath) Then
            Call InputResult.InputResult(lCurRow, "not exit")
        ElseIf sTargetFileName = sChangeFileName Then
            Call InputResult.InputResult(lCurRow, "not necessary")
        Else
            oFileSystemObject.GetFile(sTargetFilePath).Name = sChangeFileName
            Call InputResult.InputResult(lCurRow, "")
        End If
    Next lCurRow
End Sub

