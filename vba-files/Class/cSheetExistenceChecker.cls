VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSheetExistenceChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_BookExistenceChecker As cBookExistenceChecker
Private m_BookOperator As cBookOperator

' Initialize member variables.
Private Sub Class_Initialize()
    Set m_BookExistenceChecker = New cBookExistenceChecker
    Set m_BookOperator = New cBookOperator
End Sub

' Check member variables.
' Judge whether to proceed to the next process.
Private Function checkMyValue()
    checkMyValue = True
End Function

' Check whether worksheet is exist in target workbook.
' Before call this function, should set bellow member variables :
' m_sTargetFolderPath       : Target book folder path.
' m_sTargetBookName         : Target book name.
' m_sSheetName              : Target sheet name.
' If target sheet name is not exist return false.
Public Function HasSheet(ByVal vpcTarget As cPathCreator, ByVal vsSheetName As String) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    Dim bOpened As Boolean: bOpened = False
    If Not m_BookExistenceChecker.IsBookInOpen(vpcTarget) Then: bOpened = m_BookOperator.OpenBook(vpcTarget)

    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    For Each wsTarget In wbTarget.Sheets
        If wsTarget.Name = vsSheetName Then
            HasSheet = True
            GoTo LabelExit
        End If
    Next wsTarget
    HasSheet = False
LabelExit:
    If bOpened Then: Call m_BookOperator.CloseBook(vpcTarget)
End Function

' Check whether sheet is visible in target workbook.
' Before call this function, should set bellow member variables :
' m_sTargetFolderPath       : Target book folder path.
' m_sTargetBookName         : Target book name.
' m_sSheetName              : Target sheet name.
' If target sheet name is not exist return false.
Public Function IsSheetVisible(ByVal vpcTarget As cPathCreator, ByVal vsSheetName As String) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    Dim bOpened As Boolean: bOpened = False
    If Not m_BookExistenceChecker.IsBookInOpen(vpcTarget) Then: bOpened = m_BookOperator.OpenBook(vpcTarget)

    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    For Each wsTarget In wbTarget.Sheets
        If wsTarget.Name = vsSheetName Then
            IsSheetVisible = wsTarget.Visible
            GoTo LabelExit
        End If
    Next wsTarget
    IsSheetVisible = False
LabelExit:
    If bOpened Then: Call m_BookOperator.CloseBook(vpcTarget)
End Function

