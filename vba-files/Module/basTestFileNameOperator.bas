Attribute VB_Name = "basTestFileNameOperator"
Option Explicit

Private Const M_S_TEST_NAME As String = "TestFileNameOperator"
Private Const M_S_TEST_ORDER1 As String = "testGetFileCount"
Private Const M_S_TEST_ORDER2 As String = "testGetFileNameArray"
Private Const M_S_TEST_ORDER3 As String = "testGetFileName"
Private Const M_S_TEST_ORDER4 As String = "testChangeFileNameArray"
Private Const M_S_TEST_ORDER5 As String = "testChangeFileName"

Private m_pcThisTest As cPathCreator
Private m_FileNameOperator As cFileNameOperator

' Order1
Public Function DoTestGetFileCount() As Boolean
    Call init
    DoTestGetFileCount = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER1)
    Call deinit
End Function

' Order2
Public Function DoTestGetFileNameArray() As Boolean
    Call init
    DoTestGetFileNameArray = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER2)
    Call deinit
End Function

' Order3
Public Function DoTestGetFileName() As Boolean
    Call init
    DoTestGetFileName = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER3)
    Call deinit
End Function

' Order4
Public Function DoTestChangeFileNameArray() As Boolean
    Call init
    DoTestChangeFileNameArray = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER4)
    Call deinit
End Function

' Order5
Public Function DoTestChangeFileName() As Boolean
    Call init
    DoTestChangeFileName = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER5)
    Call deinit
End Function

' Initialize this test.
Private Sub init()
    Set m_pcThisTest = New cPathCreator
    Set m_FileNameOperator = New cFileNameOperator
    
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
    Set m_FileNameOperator = Nothing
End Sub

' Order1
' Verify GetFileFormatString.
Private Function testGetFileCount() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileExtension = G_S_EXTENSION_CSV
    If 5 <> m_FileNameOperator.GetFileCount(m_pcThisTest) Then: GoTo LabelFailure
    
    m_pcThisTest.sFileExtension = G_S_EXTENSION_XLSX
    If 5 <> m_FileNameOperator.GetFileCount(m_pcThisTest) Then: GoTo LabelFailure
    
    m_pcThisTest.sFileExtension = G_S_EXTENSION_XLSM
    If 5 <> m_FileNameOperator.GetFileCount(m_pcThisTest) Then: GoTo LabelFailure
    
    m_pcThisTest.sFileExtension = G_S_EXTENSION_XLSM
    If 4 <> m_FileNameOperator.GetFileCount(m_pcThisTest, "3") Then: GoTo LabelFailure
    
LabelPass:
    testGetFileCount = True
    Exit Function

LabelFailure:
    testGetFileCount = False
End Function

' Order2
' Verify GetFileFormatString.
Private Function testGetFileNameArray() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileExtension = G_S_EXTENSION_CSV
    
    Dim vFileNameArray As Variant
    Dim lCntArray As Long
    vFileNameArray = m_FileNameOperator.GetFileNameArray(m_pcThisTest)
    For lCntArray = LBound(vFileNameArray) To UBound(vFileNameArray) - 1
        If Not M_S_TEST_NAME & "_" & (lCntArray + 1) & G_S_EXTENSION_CSV _
        = vFileNameArray(lCntArray) Then
            GoTo LabelFailure
        End If
    Next lCntArray
    
    vFileNameArray = m_FileNameOperator.GetFileNameArray(m_pcThisTest, "1")
    For lCntArray = LBound(vFileNameArray) To UBound(vFileNameArray) - 1
        If Not M_S_TEST_NAME & "_" & (lCntArray + 2) & G_S_EXTENSION_CSV _
        = vFileNameArray(lCntArray) Then
            GoTo LabelFailure
        End If
    Next lCntArray
    
LabelPass:
    testGetFileNameArray = True
    Exit Function

LabelFailure:
    testGetFileNameArray = False
End Function

' Order3
' Verify GetFileFormatEnum.
Private Function testGetFileName() As Boolean
    On Error GoTo LabelFailure
    
    m_pcThisTest.sFileExtension = G_S_EXTENSION_XLSM
    
    Dim sFileName As String
    Dim lCntArray As Long
    For lCntArray = 0 To 4
        sFileName = m_FileNameOperator.GetFileName(m_pcThisTest, lCntArray)
        If Not M_S_TEST_NAME & "_" & (lCntArray + 1) & G_S_EXTENSION_XLSM = sFileName Then
            GoTo LabelFailure
        End If
    Next lCntArray
    
    For lCntArray = 0 To 3
        sFileName = m_FileNameOperator.GetFileName(m_pcThisTest, lCntArray, "1")
        If Not M_S_TEST_NAME & "_" & (lCntArray + 2) & G_S_EXTENSION_XLSM = sFileName Then
            GoTo LabelFailure
        End If
    Next lCntArray
    
LabelPass:
    testGetFileName = True
    Exit Function

LabelFailure:
    testGetFileName = False
End Function

' Order4
' Verify HasExtension.
Private Function testChangeFileNameArray() As Boolean
    On Error GoTo LabelFailure
    
    Dim sPreFileNameArray(5) As String
    Dim sPostFileNameArray(5) As String
    Dim lCntArray As Long
    For lCntArray = 0 To 4
        sPreFileNameArray(lCntArray) = M_S_TEST_NAME & "_" & (lCntArray + 1) & G_S_EXTENSION_CSV
        sPostFileNameArray(lCntArray) = M_S_TEST_NAME & "_" & (lCntArray + 1) _
        & "_" & (lCntArray + 1) & G_S_EXTENSION_CSV
    Next lCntArray
    With m_FileNameOperator
        If Not .ChangeFileNameArray(m_pcThisTest, sPreFileNameArray, sPostFileNameArray) Then
            GoTo LabelFailure
        End If
        If "" = Dir(m_pcThisTest.sFolderPath & "\" & sPostFileNameArray(0)) Then: GoTo LabelFailure
        If "" = Dir(m_pcThisTest.sFolderPath & "\" & sPostFileNameArray(1)) Then: GoTo LabelFailure
        If "" = Dir(m_pcThisTest.sFolderPath & "\" & sPostFileNameArray(2)) Then: GoTo LabelFailure
        If "" = Dir(m_pcThisTest.sFolderPath & "\" & sPostFileNameArray(3)) Then: GoTo LabelFailure
        If "" = Dir(m_pcThisTest.sFolderPath & "\" & sPostFileNameArray(4)) Then: GoTo LabelFailure
        If Not .ChangeFileNameArray(m_pcThisTest, sPostFileNameArray, sPreFileNameArray) Then
            GoTo LabelFailure
        End If
    End With
    
LabelPass:
    testChangeFileNameArray = True
    Exit Function

LabelFailure:
    testChangeFileNameArray = False
End Function

' Order4
' Verify RemoveExtension.
Private Function testChangeFileName() As Boolean
    On Error GoTo LabelFailure
    
    Dim sPreFileName As String
    Dim sPostFileName As String
    sPreFileName = M_S_TEST_NAME & "_3" & G_S_EXTENSION_XLSX
    sPostFileName = M_S_TEST_NAME & "_3_3" & G_S_EXTENSION_XLSX
    
    m_pcThisTest.sFileName = sPreFileName
    Call m_pcThisTest.SetPath
    If Not m_FileNameOperator.ChangeFileName(m_pcThisTest, sPostFileName) Then
        GoTo LabelFailure
    End If
    If "" = Dir(m_pcThisTest.sFolderPath & "\" & sPostFileName) Then: GoTo LabelFailure
    
    m_pcThisTest.sFileName = sPostFileName
    m_pcThisTest.sFullPath = ""
    Call m_pcThisTest.SetPath
    If Not m_FileNameOperator.ChangeFileName(m_pcThisTest, sPreFileName) Then
        GoTo LabelFailure
    End If
    
LabelPass:
    testChangeFileName = True
    Exit Function

LabelFailure:
    testChangeFileName = False
End Function


