VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileExistenceChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sTargetFolderPath As String
Private m_sTargetFileName As String
Private m_lTargetFileFormat As Long
Private m_sTargetSheetName As String

Public Property Get sTargetFolderPath() As String
    sTargetFolderPath = m_sTargetFolderPath
End Property
Public Property Let sTargetFolderPath(ByVal vsTargetFolderPath As String)
    m_sTargetFolderPath = vsTargetFolderPath
End Property

Public Property Get sTargetFileName() As String
    sTargetFileName = m_sTargetFileName
End Property
Public Property Let sTargetFileName(ByVal vsTargetFileName As String)
    m_sTargetFileName = vsTargetFileName
End Property

Public Property Get sTargetSheetName() As String
    sTargetSheetName = m_sTargetSheetName
End Property
Public Property Let sTargetSheetName(ByVal vsTargetSheetName As String)
    m_sTargetSheetName = vsTargetSheetName
End Property

' Initialize member variables.
Private Sub Class_Initialize()
    m_sTargetFolderPath = ""
    m_sTargetFileName = ""
    m_sTargetSheetName = ""
End Sub

' Check member variables.
' Judge whether to proceed to the next process.
Private Function checkMyValue()
    checkMyValue = True
End Function

' Check whether workbook is exist in opening workbooks.
' Before call this function, should set bellow member variables :
' m_sTargetFileName         : Target file name.
' If target workbook name is not exist, return false.
Public Function IsWorkbookInOpening() As Boolean
    Dim wbTarget As Workbook
    
    For Each wbTarget In Workbooks
        If wbTarget.Name = m_sTargetFileName Then
            IsWorkbookInOpening = True
            Exit Function
        End If
    Next wbTarget
    IsWorkbookInOpening = False
End Function

' Check whether worksheet is exist in target workbook.
' Before call this function, should set bellow member variables :
' m_sTargetFolderPath       : Target file folder path.
' m_sTargetFileName         : Target file name.
' m_sTargetSheetName        : Target worksheet name.
' If worksheet of param name is not exist return false.
Public Function HasWorksheet() As Boolean
    Dim sTargetFileFullPath As String
    Dim sTargetFileFormat As String
    Dim wbTarget As Workbook
    Dim bOpened As Boolean: bOpened = False
    
    sTargetFileFullPath = m_sTargetFolderPath & "\" & m_sTargetFileName
    
    If "" = Dir(sTargetFileFullPath) Then: Exit Function
    If Not IsWorkbookInOpening Then
        Workbooks.Open (sTargetFileFullPath)
        bOpened = True
    End If
    Set wbTarget = Workbooks(Dir(sTargetFileFullPath))

    Dim wsTarget As Worksheet
    For Each wsTarget In wbTarget.Sheets
        If wsTarget.Name = m_sTargetSheetName Then
            HasWorksheet = True
            GoTo LabelExit
        End If
    Next wsTarget
    HasWorksheet = False
LabelExit:
    If bOpened Then: wbTarget.Close
End Function

