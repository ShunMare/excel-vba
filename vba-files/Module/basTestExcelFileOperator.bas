Attribute VB_Name = "basTestExcelFileOperator"
Option Explicit

Private Enum M_E_ORDER
    ORDER_1
    ORDER_2
    ORDER_3
    ORDER_4
    ORDER_5
    ORDER_6
    ORDER_MAX
End Enum

Private Enum M_E_PATH
    FILE_VALIDATION
    FOLDER_VALIDATION
    FOLDER_TESTS
    FOLDER_THIS_TEST
    PATH_MAX
End Enum

Private Const M_S_TEST_BOOK_OPERATOR_ORDER As String = "TestBookOperatorOrder"
Private Const M_S_TEST_ORDER1 As String = "TestBookOperatorOrder1"
Private Const M_S_TEST_ORDER2 As String = "TestBookOperatorOrder2"
Private Const M_S_TEST_ORDER3 As String = "TestBookOperatorOrder3"
Private Const M_S_TEST_ORDER4 As String = "TestBookOperatorOrder4"
Private Const M_S_TEST_ORDER5 As String = "TestBookOperatorOrder5"
Private Const M_S_TEST_ORDER6 As String = "TestBookOperatorOrder6"

Private m_bResult As Boolean
Private m_BookOperator As cBookOperator
Private m_PathCreator As cPathCreator
Private m_BookExistenceChecker As cBookExistenceChecker

' Verify all test in this module
Public Function TestAll() As Boolean()
    Dim bResult(M_E_ORDER.ORDER_MAX) As Boolean
    bResult(M_E_ORDER.ORDER_1) = Order1
    bResult(M_E_ORDER.ORDER_2) = Order2
    bResult(M_E_ORDER.ORDER_3) = Order3
    bResult(M_E_ORDER.ORDER_4) = Order4
    bResult(M_E_ORDER.ORDER_5) = Order5
    bResult(M_E_ORDER.ORDER_6) = Order6
    TestAll = bResult
End Function

' Click Order1
Public Function Order1()
    Call init
    Order1 = testInitializeProperty
    Call deinit
End Function

' Click Order2
Public Function Order2()
    Call init
    Order2 = testCreateBook
    Call deinit
End Function

' Click Order3
Public Function Order3()
    Call init
    Order3 = testOpenBook
    Call deinit
End Function

' Click Order4
Public Function Order4()
    Call init
    Order4 = testCloseBook
    Call deinit
End Function

' Click Order5
Public Function Order5()
    Call init
    Order5 = testCopyBook
    Call deinit
End Function

' Click Order6
Public Function Order6()
    Call init
    Order6 = testRemoveBook
    Call deinit
End Function

' Initialize this test.
Private Sub init()
    Set m_BookOperator = New cBookOperator
    Set m_PathCreator = New cPathCreator
    Set m_BookExistenceChecker = New cBookExistenceChecker
    With m_PathCreator
    ' varidation file path
        Call .CombineFolderPath(M_E_PATH.FILE_VALIDATION, _
        ThisWorkbook.Path & "\" & G_S_TEST_FOLDER_NAME)
        Call .SetFileName(M_E_PATH.FILE_VALIDATION, _
        G_S_FILE_NAME_VALIDATION & G_S_EXTENSION_XLSX)
    ' varidation folder path
        Call .CombineFolderPath(M_E_PATH.FOLDER_VALIDATION, _
        ThisWorkbook.Path & "\" & G_S_TEST_FOLDER_NAME & "\" & G_S_TEST_VALIDATION_FOLDER_NAME)
    ' test folder path
        Call .CombineFolderPath(M_E_PATH.FOLDER_TESTS, _
        ThisWorkbook.Path & "\" & G_S_TEST_FOLDER_NAME)
    ' TestBookOperatorOrder folder path
        Call .CombineFolderPath(M_E_PATH.FOLDER_THIS_TEST, _
        ThisWorkbook.Path & "\" & G_S_TEST_FOLDER_NAME & "\" & M_S_TEST_BOOK_OPERATOR_ORDER)
    End With
    m_bResult = True
    Application.ScreenUpdating = False
End Sub

' Deinitialize this test.
Private Sub deinit()
    Application.ScreenUpdating = True
    Set m_BookOperator = Nothing
End Sub

' Order1
' Verify initialize property.
' Check member variables is initialized.
Private Function testInitializeProperty() As Boolean
    On Error GoTo LabelFailure
    testInitializeProperty = m_bResult
    Exit Function

LabelFailure:
    m_bResult = False
    testInitializeProperty = m_bResult
End Function

' Order2
' Verify CreateBook.
' Generate book in validation folder.
' Before generate file, check the file is existing.If generating file is already existing, remove the file.
' Finally check generated file is exist or not.
Private Function testCreateBook() As Boolean
    On Error GoTo LabelFailure
' Set path information.
    Call m_PathCreator.SetFileName(M_E_PATH.FOLDER_VALIDATION, M_S_TEST_ORDER2 & G_S_EXTENSION_XLSX)
    Dim pcTarget As New cPathCreator
    pcTarget.sFullPath = m_PathCreator.GetFullPath(M_E_PATH.FOLDER_VALIDATION)

' Create book.
    With m_BookOperator
        If "" <> Dir(pcTarget.sFullPath) Then: Kill (pcTarget.sFullPath)
        Call .CreateBook(pcTarget)
        If "" = Dir(pcTarget.sFullPath) Then: m_bResult = False
    End With
    testCreateBook = m_bResult
    Exit Function

LabelFailure:
    m_bResult = False
    testCreateBook = m_bResult
End Function

' Order3
' Verify CopyBook.
' Open book.
Private Function testOpenBook() As Boolean
    On Error GoTo LabelFailure
' Set path information.
    Dim pcTarget As New cPathCreator
    Call m_PathCreator.SetFileName(M_E_PATH.FOLDER_THIS_TEST, _
    M_S_TEST_ORDER3 & G_S_EXTENSION_XLSX)
    pcTarget.sFileName = m_PathCreator.GetFileName(M_E_PATH.FOLDER_THIS_TEST)
    pcTarget.sFullPath = m_PathCreator.GetFullPath(M_E_PATH.FOLDER_THIS_TEST)

' Open book
    Call m_BookOperator.OpenBook(pcTarget)
    With m_BookExistenceChecker
        Call .SetPathInfo(pcTarget)
        If Not .IsBookInOpen Then
            m_bResult = False
            testOpenBook = m_bResult
            Exit Function
        End If
    End With
    Application.DisplayAlerts = False
    Workbooks(Dir(pcTarget.sFullPath)).Close SaveChanges:=True
    Application.DisplayAlerts = True
    testOpenBook = m_bResult
    Exit Function
    
LabelFailure:
    m_bResult = False
    testOpenBook = m_bResult
End Function

' Order4
' Verify CloseBook.
Private Function testCloseBook() As Boolean
    On Error GoTo LabelFailure
' Set path information.
    Dim pcTarget As New cPathCreator
    Call m_PathCreator.SetFileName(M_E_PATH.FOLDER_THIS_TEST, _
    M_S_TEST_ORDER4 & G_S_EXTENSION_XLSX)
    pcTarget.sFileName = m_PathCreator.GetFileName(M_E_PATH.FOLDER_THIS_TEST)
    pcTarget.sFullPath = m_PathCreator.GetFullPath(M_E_PATH.FOLDER_THIS_TEST)
    
    Workbooks.Open (pcTarget.sFullPath)
    With m_BookOperator
        Call .CloseBook(pcTarget)
    End With
    With m_BookExistenceChecker
        Call .SetPathInfo(pcTarget)
        If .IsBookInOpen Then: m_bResult = False
    End With
    testCloseBook = m_bResult
    Exit Function

LabelFailure:
    m_bResult = False
    testCloseBook = m_bResult
End Function

' Order5
' Verify CopyBook.
' Copy
Private Function testCopyBook() As Boolean
    On Error GoTo LabelFailure
    Dim pcSource As New cPathCreator
    Dim pcDestination As New cPathCreator
    Call m_PathCreator.SetFileName(M_E_PATH.FOLDER_THIS_TEST, _
    M_S_TEST_ORDER5 & G_S_EXTENSION_XLSX)
    pcSource.sFileName = m_PathCreator.GetFileName(M_E_PATH.FOLDER_THIS_TEST)
    pcSource.sFullPath = m_PathCreator.GetFullPath(M_E_PATH.FOLDER_THIS_TEST)
    Call m_PathCreator.SetFileName(M_E_PATH.FOLDER_VALIDATION, _
    M_S_TEST_ORDER5 & G_S_EXTENSION_XLSX)
    pcSource.sFileName = m_PathCreator.GetFileName(M_E_PATH.FOLDER_VALIDATION)
    pcSource.sFullPath = m_PathCreator.GetFullPath(M_E_PATH.FOLDER_VALIDATION)
    
    With m_BookOperator
        Call .SetSourcePathInfo(pcSource)
        Call .SetDestinationPathInfo(pcDestination)
        Call .CopyBook
    End With
    testCopyBook = m_bResult
    Exit Function

LabelFailure:
    m_bResult = False
    testCopyBook = m_bResult
End Function

' Order6
' Verify RemoveBook.
Private Function testRemoveBook() As Boolean
    On Error GoTo LabelFailure
    Dim pcTarget As New cPathCreator
    Call m_PathCreator.SetFileName(M_E_PATH.FOLDER_THIS_TEST, _
    M_S_TEST_ORDER6 & G_S_EXTENSION_XLSX)
    pcTarget.sFileName = m_PathCreator.GetFileName(M_E_PATH.FOLDER_THIS_TEST)
    pcTarget.sFullPath = m_PathCreator.GetFullPath(M_E_PATH.FOLDER_THIS_TEST)
    
    With m_BookOperator
        Call .RemoveBook(pcTarget)
        If "" <> Dir(pcTarget.sFullPath) Then: m_bResult = False
        Call .CreateBook(pcTarget)
    End With
    testRemoveBook = m_bResult
    Exit Function

LabelFailure:
    m_bResult = False
    testRemoveBook = m_bResult
End Function

