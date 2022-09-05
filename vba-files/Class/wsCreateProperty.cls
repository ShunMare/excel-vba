VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsCreateProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const M_L_START_ROW As Long = 2
Private Const M_L_MEMBER_COL As Long = 1
Private Const M_L_FORM_COL As Long = 2
Private Const M_L_PROPERTY_COL As Long = 3
Private Const M_S_CODE_FILE_NAME As string = "CreatePropertyCode.txt"
Private Const M_S_PROPERTY_NAME As String = "{$property_name}"
Private Const M_S_PROPERTY_FORM As String = "{$property_form}"
Private Const M_S_INIT_CONTENT As String = "{$init_content}"
Private Const M_S_CHECK_CONTENT As String = "{$check_content}"
Private Const M_S_COMP_CONTENT As String = "{$comp_content}"
Private Const M_S_INIT_PART_CONTENT As String = "    " & "m_" & M_S_PROPERTY_NAME & " = " & M_S_COMP_CONTENT & vbLf
Private Const M_S_CHECK_PART_CONTENT As String = _
"    If " & M_S_COMP_CONTENT & " = m_" & M_S_PROPERTY_NAME & " Then" & vbLf & _
"        checkMyValue = False" & vbLf & _
"        Exit Function" & vbLf & _
"    End If" & vbLf
Private Const M_S_PROPERTY_FORMAT As String = _
"'--------------------------------------------------------------------------------" & vbLf & _
"'@brief : getter letter " & M_S_PROPERTY_NAME & vbLf & _
"'@param : NULL" & vbLf & _
"'@return: NULL" & vbLf & _
"'--------------------------------------------------------------------------------" & vbLf & _
"Public Property Get " & M_S_PROPERTY_NAME & "() As " & M_S_PROPERTY_FORM & vbLf & _
"    " & M_S_PROPERTY_NAME & " = m_" & M_S_PROPERTY_NAME & vbLf & _
"End Property" & vbLf & _
"" & vbLf & _
"Public Property Let " & M_S_PROPERTY_NAME & "(ByVal v" & M_S_PROPERTY_NAME & " As " & M_S_PROPERTY_FORM & ")" & vbLf & _
"    m_" & M_S_PROPERTY_NAME & " = v" & M_S_PROPERTY_NAME & vbLf & _
"End Property" & vbLf

Private Const M_S_CLASS_INIT_FORMAT As String = _
"'--------------------------------------------------------------------------------" & vbLf & _
"'@brief : initialize" & vbLf & _
"'@param : NULL" & vbLf & _
"'@return: NULL" & vbLf & _
"'--------------------------------------------------------------------------------" & vbLf & _
"Private Sub Class_initialize()" & vbLf & _
M_S_INIT_CONTENT & _
"End Sub" & vbLf

Private Const M_S_PROPERTY_CHECK_FORMAT As String = _
"'--------------------------------------------------------------------------------" & vbLf & _
"'@brief : check my value" & vbLf & _
"'@param : NULL" & vbLf & _
"'@return: True, False" & vbLf & _
"'--------------------------------------------------------------------------------" & vbLf & _
"Private Function checkMyValue()" & vbLf & _
M_S_CHECK_CONTENT & _
"    checkMyValue = True" & vbLf & _
"End Function" & vbLf

'--------------------------------------------------------------------------------
'@brief : click btnSeparateString
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnSeparateString_Click()
    Dim SeparateString As cSeparateString
    Dim lCurRow As Long
    Dim lEndRow As Long
    Dim sTarget As String
    
    lEndRow = Me.Cells(Me.Rows.Count, M_L_MEMBER_COL).End(xlUp).Row
    For lCurRow = M_L_START_ROW To lEndRow
        sTarget = Me.Cells(lCurRow, M_L_MEMBER_COL)
        If InStr(sTarget, "String") <> 0 Then
            Me.Cells(lCurRow, M_L_FORM_COL) = "String"
        ElseIf InStr(sTarget, "Long") <> 0 Then
            Me.Cells(lCurRow, M_L_FORM_COL) = "Long"
        End If
    Next lCurRow
    
    Set SeparateString = New cSeparateString
    With SeparateString
        .sWorksheetName = G_S_CREATE_PROPERTY_NAME
        .lInputCol = M_L_MEMBER_COL
        .lSeparateType = E_SEPARATE_STRING.BETWEEN
        .lStartRow = M_L_START_ROW
        .lTargetCol = M_L_MEMBER_COL
        .sFrontSeparateKey = "Dim "
        .sRearSeparateKey = " As"
        Call .SeparateString
    End With
End Sub

'--------------------------------------------------------------------------------
'@brief : click btnCreatePropertyCode
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnCreatePropertyCode_Click()
    Dim lCurRow As Long
    Dim lEndRow As Long
    Dim sPropertyName As String
    Dim sPropertyForm As String
    Dim vCompContent As Variant
    Dim sTargetCode As String
    Dim sTmpCode As String
    
    lEndRow = Me.Cells(Me.Rows.Count, M_L_MEMBER_COL).End(xlUp).Row
    Open ThisWorkbook.Path & M_S_CODE_FILE_NAME For Output As #1
    'property format
    For lCurRow = M_L_START_ROW To lEndRow
        sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "m_", "")
        sPropertyForm = Me.Cells(lCurRow, M_L_FORM_COL)
        sTargetCode = M_S_PROPERTY_FORMAT
        sTargetCode = Replace(sTargetCode, M_S_PROPERTY_NAME, sPropertyName)
        sTargetCode = Replace(sTargetCode, M_S_PROPERTY_FORM, sPropertyForm)
        Me.Cells(lCurRow, M_L_PROPERTY_COL) = sTargetCode
        Print #1, sTargetCode
    Next lCurRow
    
    'init format
    sTmpCode = ""
    For lCurRow = M_L_START_ROW To lEndRow
        sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "m_", "")
        sPropertyForm = Me.Cells(lCurRow, M_L_FORM_COL)
        If "Long" = sPropertyForm Then
            vCompContent = 0
        ElseIf "String" = sPropertyForm Then
            vCompContent = """"""
        Else
            vCompContent = """"""
        End If
        sTmpCode = sTmpCode & Replace(M_S_INIT_PART_CONTENT, M_S_PROPERTY_NAME, sPropertyName)
        sTmpCode = Replace(sTmpCode, M_S_COMP_CONTENT, vCompContent)
    Next lCurRow
    sTargetCode = M_S_CLASS_INIT_FORMAT
    sTargetCode = Replace(sTargetCode, M_S_INIT_CONTENT, sTmpCode)
    Print #1, sTargetCode
    
    'check format
    sTmpCode = ""
    For lCurRow = M_L_START_ROW To lEndRow
        sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "m_", "")
        sPropertyForm = Me.Cells(lCurRow, M_L_FORM_COL)
        If "Long" = sPropertyForm Then
            vCompContent = 0
        ElseIf "String" = sPropertyForm Then
            vCompContent = """"""
        Else
            vCompContent = """"""
        End If
        sTmpCode = sTmpCode & Replace(M_S_CHECK_PART_CONTENT, M_S_PROPERTY_NAME, sPropertyName)
        sTmpCode = Replace(sTmpCode, M_S_COMP_CONTENT, vCompContent)
    Next lCurRow
    sTargetCode = M_S_PROPERTY_CHECK_FORMAT
    sTargetCode = Replace(sTargetCode, M_S_CHECK_CONTENT, sTmpCode)
    Print #1, sTargetCode
    
    Close #1
End Sub

'--------------------------------------------------------------------------------
'@brief : click btnOpen
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnOpen_Click()
    Dim oWscriptShell As Object
    set oWscriptShell = CreateObject("Wscript.Shell")
    oWscriptShell.Run ThisWorkbook.path & M_S_CODE_FILE_NAME ,3
End Sub

'--------------------------------------------------------------------------------
'@brief : click btnClear
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub btnClear_Click()
    Dim ClearContents As New cClearContents
    Set ClearContents = New cClearContents
    
    With ClearContents
        .sWorksheetName = Me.Name
        .lStartRow = M_L_START_ROW
        .lStartCol = M_L_MEMBER_COL
        .lEndCol = M_L_PROPERTY_COL
        .lKeyCol = M_L_MEMBER_COL
        Call .ClearRangeContents
    End With
End Sub