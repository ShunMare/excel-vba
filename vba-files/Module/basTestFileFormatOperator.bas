Attribute VB_Name = "basTestFileFormatOperator"
Option Explicit

Private Const M_S_TEST_NAME As String = "TestFileFormatOperator"
Private Const M_S_TEST_ORDER1 As String = "testGetFileFormatString"
Private Const M_S_TEST_ORDER2 As String = "testGetFileFormatEnum"
Private Const M_S_TEST_ORDER3 As String = "testHasExtension"
Private Const M_S_TEST_ORDER4 As String = "testRemoveExtension"

Private m_FileFormatOperator As cFileFormatOperator

' Order1
Public Function DoTestGetFileFormatString() As Boolean
    Call init
    DoTestGetFileFormatString = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER1)
    Call deinit
End Function

' Order2
Public Function DoTestGetFileFormatEnum() As Boolean
    Call init
    DoTestGetFileFormatEnum = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER2)
    Call deinit
End Function

' Order3
Public Function DoTestHasExtension() As Boolean
    Call init
    DoTestHasExtension = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER3)
    Call deinit
End Function

' Order4
Public Function DoTestRemoveExtension() As Boolean
    Call init
    DoTestRemoveExtension = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER4)
    Call deinit
End Function

' Initialize this test.
Private Sub init()
    Set m_FileFormatOperator = New cFileFormatOperator
    Application.ScreenUpdating = False
End Sub

' Deinitialize this test.
Private Sub deinit()
    Application.ScreenUpdating = True
    Set m_FileFormatOperator = Nothing
End Sub

' Order1
' Verify GetFileFormatString.
Private Function testGetFileFormatString() As Boolean
    On Error GoTo LabelFailure
    
    With m_FileFormatOperator
        If G_S_EXTENSION_XLS <> .GetFileFormatString(xlWorkbookNormal) Then: GoTo LabelFailure
        If G_S_EXTENSION_CSV <> .GetFileFormatString(xlCSV) Then: GoTo LabelFailure
        If G_S_EXTENSION_XML <> .GetFileFormatString(xlXMLSpreadsheet) Then: GoTo LabelFailure
        If G_S_EXTENSION_XLSX <> .GetFileFormatString(xlOpenXMLWorkbook) Then: GoTo LabelFailure
        If G_S_EXTENSION_XLSM <> .GetFileFormatString(xlOpenXMLWorkbookMacroEnabled) Then: GoTo LabelFailure
        If G_S_EXTENSION_NONE <> .GetFileFormatString(xlNone) Then: GoTo LabelFailure
    
        If G_S_EXTENSION_XLS <> .GetFileFormatString(M_S_TEST_ORDER1 & G_S_EXTENSION_XLS) Then: GoTo LabelFailure
        If G_S_EXTENSION_CSV <> .GetFileFormatString(M_S_TEST_ORDER1 & G_S_EXTENSION_CSV) Then: GoTo LabelFailure
        If G_S_EXTENSION_XML <> .GetFileFormatString(M_S_TEST_ORDER1 & G_S_EXTENSION_XML) Then: GoTo LabelFailure
        If G_S_EXTENSION_XLSX <> .GetFileFormatString(M_S_TEST_ORDER1 & G_S_EXTENSION_XLSX) Then: GoTo LabelFailure
        If G_S_EXTENSION_XLSM <> .GetFileFormatString(M_S_TEST_ORDER1 & G_S_EXTENSION_XLSM) Then: GoTo LabelFailure
        If G_S_EXTENSION_NONE <> .GetFileFormatString(M_S_TEST_ORDER1 & G_S_EXTENSION_NONE) Then: GoTo LabelFailure
    End With
    
LabelPass:
    testGetFileFormatString = True
    Exit Function

LabelFailure:
    testGetFileFormatString = False
End Function

' Order2
' Verify GetFileFormatEnum.
Private Function testGetFileFormatEnum() As Boolean
    On Error GoTo LabelFailure
    
    With m_FileFormatOperator
        If xlWorkbookNormal <> .GetFileFormatEnum(G_S_EXTENSION_XLS) Then: GoTo LabelFailure
        If xlCSV <> .GetFileFormatEnum(G_S_EXTENSION_CSV) Then: GoTo LabelFailure
        If xlXMLSpreadsheet <> .GetFileFormatEnum(G_S_EXTENSION_XML) Then: GoTo LabelFailure
        If xlOpenXMLWorkbook <> .GetFileFormatEnum(G_S_EXTENSION_XLSX) Then: GoTo LabelFailure
        If xlOpenXMLWorkbookMacroEnabled <> .GetFileFormatEnum(G_S_EXTENSION_XLSM) Then: GoTo LabelFailure
        If xlNone <> .GetFileFormatEnum(G_S_EXTENSION_NONE) Then: GoTo LabelFailure
    End With
    
LabelPass:
    testGetFileFormatEnum = True
    Exit Function
    
LabelFailure:
    testGetFileFormatEnum = False
End Function

' Order3
' Verify HasExtension.
Private Function testHasExtension() As Boolean
    On Error GoTo LabelFailure
    
    With m_FileFormatOperator
        If Not .HasExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XLS) Then: GoTo LabelFailure
        If Not .HasExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_CSV) Then: GoTo LabelFailure
        If Not .HasExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XML) Then: GoTo LabelFailure
        If Not .HasExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XLSX) Then: GoTo LabelFailure
        If Not .HasExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XLSM) Then: GoTo LabelFailure
        If .HasExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_NONE) Then: GoTo LabelFailure
        If .HasExtension(M_S_TEST_ORDER3 & ".abc") Then: GoTo LabelFailure
    End With

LabelPass:
    testHasExtension = True
    Exit Function

LabelFailure:
    testHasExtension = False
End Function

' Order4
' Verify RemoveExtension.
Private Function testRemoveExtension() As Boolean
    On Error GoTo LabelFailure
    
    With m_FileFormatOperator
        If M_S_TEST_ORDER3 <> .RemoveExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XLS) Then
            GoTo LabelFailure
        End If
        If M_S_TEST_ORDER3 <> .RemoveExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_CSV) Then
            GoTo LabelFailure
        End If
        If M_S_TEST_ORDER3 <> .RemoveExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XML) Then
            GoTo LabelFailure
        End If
        If M_S_TEST_ORDER3 <> .RemoveExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XLSX) Then
            GoTo LabelFailure
        End If
        If M_S_TEST_ORDER3 <> .RemoveExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_XLSM) Then
            GoTo LabelFailure
        End If
        If M_S_TEST_ORDER3 <> .RemoveExtension(M_S_TEST_ORDER3 & G_S_EXTENSION_NONE) Then
            GoTo LabelFailure
        End If
        If M_S_TEST_ORDER3 & ".abc" <> .RemoveExtension(M_S_TEST_ORDER3 & ".abc") Then
            GoTo LabelFailure
        End If
    End With
    
LabelPass:
    testRemoveExtension = True
    Exit Function

LabelFailure:
    testRemoveExtension = False
End Function
