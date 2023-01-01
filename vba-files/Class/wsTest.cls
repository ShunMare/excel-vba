VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Enum M_E_ORDER
    ORDER_1
    ORDER_2
    ORDER_3
    ORDER_4
    ORDER_5
    ORDER_6
    ORDER_7
    ORDER_8
    ORDER_MAX
End Enum

Private Const M_S_CLASS_NAME_BOOK_OPERATOR As String = "BookOperator"
Private Const M_S_CLASS_NAME_SHEET_OPERATOR As String = "SheetOperator"
Private Const M_S_PASS As String = "PASS"
Private Const M_S_FAILURE As String = "FAILURE"
Private Const M_L_CLASS_NAME_COL As Long = 3
Private Const M_L_TEST_NAME_COL As Long = 4
Private Const M_L_RESULT_COL As Long = 8
Private Const M_L_TIME_COL As Long = 9
Private m_vClassName() As Variant

Public Sub DoTest()
    Dim sClassName As String
    Dim lClassNameRow As Long
    Dim bTestAll As Boolean
    Dim lTargetOrder As Long
    Dim lOrderRow As Long
    Dim lCurOrder As Long
    Dim bResultAll() As Boolean
    Dim bResult As Boolean
    
    lOrderRow = Me.Shapes(Application.Caller).TopLeftCell.Row
    If "" <> Me.Cells(lOrderRow, M_L_CLASS_NAME_COL) Then
        lClassNameRow = lOrderRow
        bTestAll = True
    Else
        lClassNameRow = Me.Cells(lOrderRow, M_L_CLASS_NAME_COL).End(xlUp).Row
        bTestAll = False
    End If
    sClassName = Me.Cells(lClassNameRow, M_L_CLASS_NAME_COL)
    
    Dim lCurRow As Long
    lCurRow = IIf(bTestAll, lClassNameRow + 1, lOrderRow)
    Do
        If "" = Me.Cells(lCurRow, M_L_TEST_NAME_COL) Then: Exit Do
        Call getTestResult(sClassName, lCurRow)
        If Not bTestAll Then: Exit Do
        lCurRow = lCurRow + 1
    Loop
End Sub

Private Sub init()
    m_vClassName() = Array( _
        M_S_CLASS_NAME_BOOK_OPERATOR, _
        M_S_CLASS_NAME_SHEET_OPERATOR _
    )
End Sub

Private Sub getTestResult(ByVal vsClassName As String, ByVal vlRow As Long)
    Dim sTestName As String
    Dim sTestStr As String
    Dim bResult As Boolean
    sTestName = Me.Cells(vlRow, M_L_TEST_NAME_COL)
    sTestStr = "basTest" & vsClassName & ".DoTest" & sTestName
    bResult = Application.Run(sTestStr)
    Me.Cells(vlRow, M_L_RESULT_COL) = IIf(bResult, M_S_PASS, M_S_FAILURE)
    Me.Cells(vlRow, M_L_TIME_COL) = Format(Now(), "yyyy/mm/dd hh:mm")
End Sub



