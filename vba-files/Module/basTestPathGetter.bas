Attribute VB_Name = "basTestPathGetter"
Option Explicit

Private Const M_S_TEST_NAME As String = "TestPathGetter"
Private Const M_S_TEST_ORDER1 As String = "testGetFolderPathDialog"
Private Const M_S_TEST_ORDER2 As String = "testGetDrive"

Private m_PathGetter As cPathGetter

' Order1
Public Function DoTestGetFolderPathDialog() As Boolean
    Call init
    DoTestGetFolderPathDialog = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER1)
    Call deinit
End Function

' Order2
Public Function DoTestGetDrive() As Boolean
    Call init
    DoTestGetDrive = _
    Application.Run("bas" & M_S_TEST_NAME & "." & M_S_TEST_ORDER2)
    Call deinit
End Function

' Initialize this test.
Private Sub init()
    Set m_PathGetter = New cPathGetter
    
    Application.ScreenUpdating = False
End Sub

' Deinitialize this test.
Private Sub deinit()
    Application.ScreenUpdating = True

    Set m_PathGetter = Nothing
End Sub

' Order1
' Verify CreateSheet.
Private Function testGetFolderPathDialog() As Boolean
    On Error GoTo LabelFailure
        
        MsgBox "First send key button of Esc, second select folder and click OK.", vbInformation
        
        If "" <> m_PathGetter.GetFolderPathDialog Then: GoTo LabelFailure
        If "" = Dir(m_PathGetter.GetFolderPathDialog, vbDirectory) Then: GoTo LabelFailure
        
LabelPass:
    testGetFolderPathDialog = True
    Exit Function

LabelFailure:
    testGetFolderPathDialog = False
End Function

' Order2
' Verify ShowSheet.
Private Function testGetDrive() As Boolean
    On Error GoTo LabelFailure
    
    Dim sDrive() As String
    sDrive = m_PathGetter.GetDrive
    If 0 = InStr(sDrive(0), "C") Then: GoTo LabelFailure
    
LabelPass:
    testGetDrive = True
    Exit Function

LabelFailure:
    testGetDrive = False
End Function
