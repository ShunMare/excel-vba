Attribute VB_Name = "basTestSheetOperator"
Option Explicit

Private Const M_S_TEST_NAME As String = "TestSheetOperator"
Private Const M_S_TEST_ORDER1 As String = "testCreateSheet"
Private Const M_S_TEST_ORDER2 As String = "testShowSheet"
Private Const M_S_TEST_ORDER3 As String = "testHideSheet"
Private Const M_S_TEST_ORDER4 As String = "testCopySheet"
Private Const M_S_TEST_ORDER5 As String = "testChangeSheetName"
Private Const M_S_TEST_ORDER6 As String = "testRemoveSheet"

Private m_pcValidationFile As cPathCreator
Private m_pcValidation As cPathCreator
Private m_pcTests As cPathCreator
Private m_pcThisTest As cPathCreator
Private m_BookOperator As cBookOperator
Private m_SheetOperator As cSheetOperator
Private m_BookExistenceChecker As cBookExistenceChecker
Private m_SheetExistenceChecker As cSheetExistenceChecker

' Order1
Public Function DoTestCreateSheet() As Boolean
    Call init
    DoTestCreateSheet = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER1)
    Call deinit
End Function

' Order2
Public Function DoTestShowSheet() As Boolean
    Call init
    DoTestShowSheet = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER2)
    Call deinit
End Function

' Order3
Public Function DoTestHideSheet() As Boolean
    Call init
    DoTestHideSheet = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER3)
    Call deinit
End Function

' Order4
Public Function DoTestCopySheet() As Boolean
    Call init
    DoTestCopySheet = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER4)
    Call deinit
End Function

' Order5
Public Function DoTestChangeSheetName() As Boolean
    Call init
    DoTestChangeSheetName = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER5)
    Call deinit
End Function

' Order6
Public Function DoTestRemoveSheet() As Boolean
    Call init
    DoTestRemoveSheet = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER6)
    Call deinit
End Function

' Initialize this test.
Private Sub init()
    Set m_pcValidationFile = New cPathCreator
    Set m_pcValidation = New cPathCreator
    Set m_pcTests = New cPathCreator
    Set m_pcThisTest = New cPathCreator
    Set m_BookOperator = New cBookOperator
    Set m_SheetOperator = New cSheetOperator
    Set m_BookExistenceChecker = New cBookExistenceChecker
    Set m_SheetExistenceChecker = New cSheetExistenceChecker
    
    With m_pcValidationFile
        .sFolderPath = ThisWorkbook.Path
        Call .CombineFolderPath(G_S_TEST_FOLDER_NAME)
        .sFileName = G_S_FILE_NAME_VALIDATION & G_S_EXTENSION_XLSX
        Call .SetPath
        .bLock = True
    End With
    With m_pcValidation
        .sFolderPath = ThisWorkbook.Path
        Call .CombineFolderPath(G_S_TEST_FOLDER_NAME)
        Call .CombineFolderPath(G_S_TEST_VALIDATION_FOLDER_NAME)
    End With
    With m_pcTests
        .sFolderPath = ThisWorkbook.Path
        Call .CombineFolderPath(G_S_TEST_FOLDER_NAME)
    End With
    With m_pcThisTest
        .sFolderPath = ThisWorkbook.Path
        Call .CombineFolderPath(G_S_TEST_FOLDER_NAME)
        Call .CombineFolderPath(M_S_TEST_NAME)
    End With
    
    Application.ScreenUpdating = False
End Sub

' Deinitialize this test.
Private Sub deinit()
    Application.ScreenUpdating = True
    Set m_SheetOperator = Nothing
End Sub

' Order1
' Verify CreateSheet.
' Generate book in validation folder.
' Before generate file, check the file is existing.If generating file is already existing, remove the file.
' Finally check generated file is exist or not.
Private Function testCreateSheet() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER1 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
        
    If m_SheetExistenceChecker.HasSheet(m_pcThisTest, M_S_TEST_ORDER1) Then
        Call m_BookOperator.OpenBook(m_pcThisTest)
        Application.DisplayAlerts = False
        Workbooks(m_pcThisTest.sFileName).Worksheets(M_S_TEST_ORDER1).Delete
        Application.DisplayAlerts = True
        Call m_BookOperator.CloseBook(m_pcThisTest)
    End If
    If Not m_SheetOperator.CreateSheet(m_pcThisTest, M_S_TEST_ORDER1) Then: GoTo LabelFailure
    If Not m_SheetExistenceChecker.HasSheet(m_pcThisTest, M_S_TEST_ORDER1) Then: GoTo LabelFailure
    
LabelPass:
    testCreateSheet = True
    Exit Function

LabelFailure:
    testCreateSheet = False
End Function

' Order2
' Verify ShowSheet.
Private Function testShowSheet() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER2 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
    
    If m_SheetExistenceChecker.IsSheetVisible(m_pcThisTest, M_S_TEST_ORDER2) Then
        Call m_BookOperator.OpenBook(m_pcThisTest)
        Workbooks(m_pcThisTest.sFileName).Worksheets(M_S_TEST_ORDER2).Visible = False
        Call m_BookOperator.CloseBook(m_pcThisTest)
    End If
    If Not m_SheetOperator.ShowSheet(m_pcThisTest, M_S_TEST_ORDER2) Then: GoTo LabelFailure
    If Not m_SheetExistenceChecker.IsSheetVisible(m_pcThisTest, M_S_TEST_ORDER2) Then: GoTo LabelFailure
    
LabelPass:
    testShowSheet = True
    Exit Function

LabelFailure:
    testShowSheet = False
End Function

' Order3
' Verify HideSheet.
Private Function testHideSheet() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER3 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
        
    If Not m_SheetExistenceChecker.IsSheetVisible(m_pcThisTest, M_S_TEST_ORDER3) Then
        Call m_BookOperator.OpenBook(m_pcThisTest)
        Workbooks(m_pcThisTest.sFileName).Worksheets(M_S_TEST_ORDER3).Visible = True
        Call m_BookOperator.CloseBook(m_pcThisTest)
    End If
    If Not m_SheetOperator.HideSheet(m_pcThisTest, M_S_TEST_ORDER3) Then: GoTo LabelFailure
    If m_SheetExistenceChecker.IsSheetVisible(m_pcThisTest, M_S_TEST_ORDER3) Then: GoTo LabelFailure

LabelPass:
    testHideSheet = True
    Exit Function

LabelFailure:
    testHideSheet = False
End Function

' Order4
' Verify CopySheet.
Private Function testCopySheet() As Boolean
    On Error GoTo LabelFailure
    
    Dim pcSource As New cPathCreator
    Set pcSource = m_pcValidationFile
    
    Dim pcDestination As New cPathCreator
    pcDestination.sFolderPath = m_pcValidation.sFolderPath
    pcDestination.sFileName = M_S_TEST_ORDER4 & G_S_EXTENSION_XLSX
    Call pcDestination.SetPath
        
    If m_SheetExistenceChecker.HasSheet(pcDestination, G_S_SHEET_NAME_VALIDATION4) Then
        Call m_BookOperator.OpenBook(pcDestination)
        Application.DisplayAlerts = False
        Workbooks(pcDestination.sFileName).Worksheets(G_S_SHEET_NAME_VALIDATION4).Delete
        Application.DisplayAlerts = True
        Call m_BookOperator.CloseBook(pcDestination)
    End If
    If Not m_SheetOperator.CopySheet(pcSource, pcDestination, G_S_SHEET_NAME_VALIDATION4) Then
        GoTo LabelFailure
    End If
    If Not m_SheetExistenceChecker.HasSheet(pcDestination, G_S_SHEET_NAME_VALIDATION4) Then
        GoTo LabelFailure
    End If

LabelPass:
    testCopySheet = True
    Exit Function

LabelFailure:
    testCopySheet = False
End Function

' Order5
' Verify CopySheet.
Private Function testChangeSheetName() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER5 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
        
    If Not m_SheetExistenceChecker.HasSheet(m_pcThisTest, G_S_SHEET_NAME_VALIDATION5) Then
        Call m_BookOperator.OpenBook(m_pcThisTest)
        Application.DisplayAlerts = False
        Workbooks(m_pcThisTest.sFileName).Worksheets(1).Name = G_S_SHEET_NAME_VALIDATION5
        Application.DisplayAlerts = True
        Call m_BookOperator.CloseBook(m_pcThisTest)
    End If
    If Not m_SheetOperator.ChangeSheetName(m_pcThisTest, G_S_SHEET_NAME_VALIDATION5, M_S_TEST_ORDER5) Then
        GoTo LabelFailure
    End If
    If Not m_SheetExistenceChecker.HasSheet(m_pcThisTest, M_S_TEST_ORDER5) Then
        GoTo LabelFailure
    End If

LabelPass:
    testChangeSheetName = True
    Exit Function

LabelFailure:
    testChangeSheetName = False
End Function

' Order6
' Verify RemoveBook.
Private Function testRemoveSheet() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER6 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
    
    If Not m_SheetExistenceChecker.IsSheetVisible(m_pcThisTest, M_S_TEST_ORDER6) Then
        Call m_BookOperator.OpenBook(m_pcThisTest)
        Workbooks(m_pcThisTest.sFileName).Worksheets.Add before:=Worksheets(1)
        Workbooks(m_pcThisTest.sFileName).Worksheets(1).Name = M_S_TEST_ORDER6
        Call m_BookOperator.CloseBook(m_pcThisTest)
    End If
    If Not m_SheetOperator.RemoveSheet(m_pcThisTest, M_S_TEST_ORDER6) Then: GoTo LabelFailure
    If m_SheetExistenceChecker.HasSheet(m_pcThisTest, M_S_TEST_ORDER6) Then: GoTo LabelFailure

LabelPass:
    testRemoveSheet = True
    Exit Function

LabelFailure:
    testRemoveSheet = False
End Function
