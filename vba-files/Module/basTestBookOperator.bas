Attribute VB_Name = "basTestBookOperator"
Option Explicit

Private Const M_S_TEST_NAME As String = "TestBookOperator"
Private Const M_S_TEST_ORDER1 As String = "testCreateBook"
Private Const M_S_TEST_ORDER2 As String = "testOpenBook"
Private Const M_S_TEST_ORDER3 As String = "testCloseBook"
Private Const M_S_TEST_ORDER4 As String = "testCopyBook"
Private Const M_S_TEST_ORDER5 As String = "testRemoveBook"

Private m_pcValidationFile As cPathCreator
Private m_pcValidation As cPathCreator
Private m_pcTests As cPathCreator
Private m_pcThisTest As cPathCreator
Private m_BookOperator As cBookOperator
Private m_BookExistenceChecker As cBookExistenceChecker

' Order1
Public Function DoTestCreateBook() As Boolean
    Call init
    DoTestCreateBook = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER1)
    Call deinit
End Function

' Order2
Public Function DoTestOpenBook() As Boolean
    Call init
    DoTestOpenBook = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER2)
    Call deinit
End Function

' Order3
Public Function DoTestCloseBook() As Boolean
    Call init
    DoTestCloseBook = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER3)
    Call deinit
End Function

' Order4
Public Function DoTestCopyBook() As Boolean
    Call init
    DoTestCopyBook = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER4)
    Call deinit
End Function

' Order5
Public Function DoTestRemoveBook() As Boolean
    Call init
    DoTestRemoveBook = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER5)
    Call deinit
End Function

' Initialize this test.
Private Sub init()
    Set m_pcValidationFile = New cPathCreator
    Set m_pcValidation = New cPathCreator
    Set m_pcTests = New cPathCreator
    Set m_pcThisTest = New cPathCreator
    Set m_BookOperator = New cBookOperator
    Set m_BookExistenceChecker = New cBookExistenceChecker
    
    With m_pcValidationFile
        .sFolderPath = ThisWorkbook.Path
        Call .CombineFolderPath(G_S_TEST_FOLDER_NAME)
        .sFileName = G_S_FILE_NAME_VALIDATION
        .sFileExtension = G_S_EXTENSION_XLSX
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
    
    Set m_pcValidationFile = Nothing
    Set m_pcValidation = Nothing
    Set m_pcTests = Nothing
    Set m_pcThisTest = Nothing
    Set m_BookOperator = Nothing
    Set m_BookExistenceChecker = Nothing
End Sub

' Order1
' Verify CreateBook.
' Generate book in validation folder.
' Before generate file, check the file is existing.If generating file is already existing, remove the file.
' Finally check generated file is exist or not.
Private Function testCreateBook() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER1 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath

    If "" <> Dir(m_pcThisTest.sFullPath) Then: Kill (m_pcThisTest.sFullPath)
    If Not m_BookOperator.CreateBook(m_pcThisTest) Then: GoTo LabelFailure
    If "" = Dir(m_pcThisTest.sFullPath) Then: GoTo LabelFailure

LabelPass:
    testCreateBook = True
    Exit Function

LabelFailure:
    testCreateBook = False
End Function

' Order2
' Verify OpenBook.
Private Function testOpenBook() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER2 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
    
    If Not m_BookOperator.OpenBook(m_pcThisTest) Then: GoTo LabelFailure
    If Not m_BookExistenceChecker.IsBookInOpen(m_pcThisTest) Then: GoTo LabelFailure
    
    Application.DisplayAlerts = False
    Workbooks(m_pcThisTest.sFileName).Close SaveChanges:=True
    Application.DisplayAlerts = True
    
LabelPass:
    testOpenBook = True
    Exit Function
    
LabelFailure:
    testOpenBook = False
End Function

' Order3
' Verify CloseBook.
Private Function testCloseBook() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER3 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
    
    Workbooks.Open (m_pcThisTest.sFullPath)
    If Not m_BookOperator.CloseBook(m_pcThisTest) Then: GoTo LabelFailure
    If m_BookExistenceChecker.IsBookInOpen(m_pcThisTest) Then: GoTo LabelFailure

LabelPass:
    testCloseBook = True
    Exit Function

LabelFailure:
    testCloseBook = False
End Function

' Order4
' Verify CopyBook.
Private Function testCopyBook() As Boolean
    On Error GoTo LabelFailure
    
    Dim pcSource As New cPathCreator
    pcSource.sFolderPath = m_pcThisTest.sFolderPath
    pcSource.sFileName = M_S_TEST_ORDER4 & G_S_EXTENSION_XLSX
    Call pcSource.SetPath
    
    Dim pcDestination As New cPathCreator
    pcDestination.sFolderPath = m_pcValidation.sFolderPath
    pcDestination.sFileName = M_S_TEST_ORDER4 & G_S_EXTENSION_XLSX
    Call pcDestination.SetPath
    
    Application.DisplayAlerts = False
    If "" <> Dir(pcDestination.sFullPath) Then: Kill (pcDestination.sFullPath)
    Application.DisplayAlerts = True
    If Not m_BookOperator.CopyBook(pcSource, pcDestination) Then: GoTo LabelFailure
    If "" = Dir(pcDestination.sFullPath) Then: GoTo LabelFailure
    
LabelPass:
    testCopyBook = True
    Exit Function

LabelFailure:
    testCopyBook = False
End Function

' Order5
' Verify RemoveBook.
Private Function testRemoveBook() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileName = M_S_TEST_ORDER5 & G_S_EXTENSION_XLSX
    Call m_pcThisTest.SetPath
    
    If "" = Dir(m_pcThisTest.sFullPath) Then: Call m_BookOperator.CreateBook(m_pcThisTest)
    If Not m_BookOperator.RemoveBook(m_pcThisTest) Then: GoTo LabelFailure
    If "" <> Dir(m_pcThisTest.sFullPath) Then: GoTo LabelFailure

LabelPass:
    testRemoveBook = True
    Exit Function

LabelFailure:
    testRemoveBook = False
End Function
