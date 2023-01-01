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

' Private Const M_L_START_ROW As Long = 2
' Private Const M_L_MEMBER_COL As Long = 1
' Private Const M_L_FORM_COL As Long = 2
' Private Const M_L_PROPERTY_COL As Long = 3
' Private Const M_S_CODE_FILE_NAME As String = "CreatePropertyCode.txt"
' Private Const M_S_PROPERTY_NAME As String = "{$property_name}"
' Private Const M_S_PROPERTY_FORM As String = "{$property_form}"
' Private Const M_S_INIT_CONTENT As String = "{$init_content}"
' Private Const M_S_CHECK_CONTENT As String = "{$check_content}"
' Private Const M_S_COMP_CONTENT As String = "{$comp_content}"
' ''VBA format
' Private Const M_S_VBA_INIT_PART_CONTENT As String = "    " & "m_" & M_S_PROPERTY_NAME & " = " & M_S_COMP_CONTENT & vbLf
' Private Const M_S_VBA_CHECK_PART_CONTENT As String = _
' "    If " & M_S_COMP_CONTENT & " = m_" & M_S_PROPERTY_NAME & " Then" & vbLf & _
' "        checkMyValue = False" & vbLf & _
' "        Exit Function" & vbLf & _
' "    End If" & vbLf
' Private Const M_S_VBA_PROPERTY_FORMAT As String = _
' "'--------------------------------------------------------------------------------" & vbLf & _
' "'@brief : getter letter " & M_S_PROPERTY_NAME & vbLf & _
' "'@param : NULL" & vbLf & _
' "'@return: NULL" & vbLf & _
' "'--------------------------------------------------------------------------------" & vbLf & _
' "Public Property Get " & M_S_PROPERTY_NAME & "() As " & M_S_PROPERTY_FORM & vbLf & _
' "    " & M_S_PROPERTY_NAME & " = m_" & M_S_PROPERTY_NAME & vbLf & _
' "End Property" & vbLf & _
' "" & vbLf & _
' "Public Property Let " & M_S_PROPERTY_NAME & "(ByVal v" & M_S_PROPERTY_NAME & " As " & M_S_PROPERTY_FORM & ")" & vbLf & _
' "    m_" & M_S_PROPERTY_NAME & " = v" & M_S_PROPERTY_NAME & vbLf & _
' "End Property" & vbLf

' Private Const M_S_VBA_CLASS_INIT_FORMAT As String = _
' "'--------------------------------------------------------------------------------" & vbLf & _
' "'@brief : initialize" & vbLf & _
' "'@param : NULL" & vbLf & _
' "'@return: NULL" & vbLf & _
' "'--------------------------------------------------------------------------------" & vbLf & _
' "Private Sub Class_initialize()" & vbLf & _
' M_S_INIT_CONTENT & _
' "End Sub" & vbLf

' Private Const M_S_VBA_PROPERTY_CHECK_FORMAT As String = _
' "'--------------------------------------------------------------------------------" & vbLf & _
' "'@brief : check my value" & vbLf & _
' "'@param : NULL" & vbLf & _
' "'@return: True, False" & vbLf & _
' "'--------------------------------------------------------------------------------" & vbLf & _
' "Private Function checkMyValue()" & vbLf & _
' M_S_CHECK_CONTENT & _
' "    checkMyValue = True" & vbLf & _
' "End Function" & vbLf

' ''GAS format
' Private Const M_S_GAS_INIT_PART_CONTENT As String = ""
' Private Const M_S_GAS_CHECK_PART_CONTENT As String = _
' "    If " & M_S_COMP_CONTENT & " = m_" & M_S_PROPERTY_NAME & " Then" & vbLf & _
' "        checkMyValue = False" & vbLf & _
' "        Exit Function" & vbLf & _
' "    End If" & vbLf
' Private Const M_S_GAS_PROPERTY_FORMAT As String = _
' "  get " & M_S_PROPERTY_NAME & "(): " & M_S_PROPERTY_FORM & "{" & vbLf & _
' "    return this._" & M_S_PROPERTY_NAME & ";" & vbLf & _
' "  }" & vbLf & _
' "  set " & M_S_PROPERTY_NAME & "(" & M_S_PROPERTY_NAME & ": " & M_S_PROPERTY_FORM & "){" & vbLf & _
' "    this._" & M_S_PROPERTY_NAME & " = " & M_S_PROPERTY_NAME & ";" & vbLf & _
' "  }" & vbLf & vbLf

' Private Const M_S_GAS_CLASS_INIT_FORMAT As String = ""
' Private Const M_S_GAS_PROPERTY_CHECK_FORMAT As String = ""

' Private m_siCreateProperty As cSheetInfo

' '--------------------------------------------------------------------------------
' '@brief : init
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Private Sub init()
'     Set m_siCreateProperty = New cSheetInfo
'     With m_siCreateProperty
'         .sWorksheetName = G_S_CNG_FILE_NAME
'         Call .SetSheetInfo
'         .lStartRow = M_L_START_ROW
'         .lStartCol = M_L_MEMBER_COL
'         .lEndCol = M_L_PROPERTY_COL
'         .lKeyCol = M_L_MEMBER_COL
'         Call .SetRowAndColInfo
'     End With
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnSeparateString
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnSeparateString_Click()
'     Dim WordReplacer    As cWordReplacer
'     Dim WordSeparater   As cWordSeparater
'     Dim lCurRow         As Long
'     Dim lEndRow         As Long
'     Dim sTarget         As String
    
'     Call init
    
'     lEndRow = Me.Cells(Me.Rows.Count, M_L_MEMBER_COL).End(xlUp).Row
'     For lCurRow = M_L_START_ROW To lEndRow
'         sTarget = Me.Cells(lCurRow, M_L_MEMBER_COL)
'         If InStr(sTarget, "String") <> 0 Or InStr(sTarget, "string") <> 0 Then
'             Me.Cells(lCurRow, M_L_FORM_COL) = "string"
'         ElseIf InStr(sTarget, "Long") <> 0 Then
'             Me.Cells(lCurRow, M_L_FORM_COL) = "long"
'         ElseIf InStr(sTarget, "Boolean") <> 0 Or InStr(sTarget, "boolean") <> 0 Then
'             Me.Cells(lCurRow, M_L_FORM_COL) = "boolean"
'         ElseIf InStr(sTarget, "number") <> 0 Then
'             Me.Cells(lCurRow, M_L_FORM_COL) = "number"
'         End If
'     Next lCurRow
    
'     Set WordReplacer = New cWordReplacer
'     With WordReplacer
'         .sReplaceWord = ""
'         .sTargetWord = "?"
'         .replaceWords (m_siCreateProperty)
'         .sTargetWord = " "
'         .replaceWords (m_siCreateProperty)
'     End With
    
'     Set WordSeparater = New cWordSeparater
'     With WordSeparater
'         .lSeparateType = E_SEPARATE_STRING.BETWEEN
'         If Me.Range(G_S_CELL_NAME_TARHET_PROGRAM_LANGUAGE) = G_S_VBA Then
'             .sFrontSeparateKey = "Dim"
'             .sRearSeparateKey = "As"
'         ElseIf Me.Range(G_S_CELL_NAME_TARHET_PROGRAM_LANGUAGE) = G_S_GAS Then
'             .sFrontSeparateKey = "private"
'             .sRearSeparateKey = ":"
'         End If
'         Call .SeparateString(m_siCreateProperty)
'     End With
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnCreatePropertyCode
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnCreatePropertyCode_Click()
'     Dim lCurRow As Long
'     Dim lEndRow As Long
'     Dim sPropertyName As String
'     Dim sPropertyForm As String
'     Dim vCompContent As Variant
'     Dim sTargetCode As String
'     Dim sTmpCode As String
    
'     Call init
    
'     lEndRow = m_siCreateProperty.lEndRow
'     Open ThisWorkbook.Path & M_S_CODE_FILE_NAME For Output As #1
'     'property format
'     For lCurRow = M_L_START_ROW To lEndRow
'         If Me.Range(G_S_CELL_NAME_TARHET_PROGRAM_LANGUAGE) = G_S_VBA Then
'             sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "m_", "")
'             sTargetCode = M_S_VBA_PROPERTY_FORMAT
'         ElseIf Me.Range(G_S_CELL_NAME_TARHET_PROGRAM_LANGUAGE) = G_S_GAS Then
'             sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "_", "")
'             sTargetCode = M_S_GAS_PROPERTY_FORMAT
'         End If
'         sPropertyForm = Me.Cells(lCurRow, M_L_FORM_COL)
'         sTargetCode = Replace(sTargetCode, M_S_PROPERTY_NAME, sPropertyName)
'         sTargetCode = Replace(sTargetCode, M_S_PROPERTY_FORM, sPropertyForm)
'         Me.Cells(lCurRow, M_L_PROPERTY_COL) = sTargetCode
'         Print #1, sTargetCode
'     Next lCurRow
    
'     'init format
'     If Me.Range(G_S_CELL_NAME_TARHET_PROGRAM_LANGUAGE) = G_S_VBA Then
'         sTmpCode = ""
'         For lCurRow = M_L_START_ROW To lEndRow
'             sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "m_", "")
'             sPropertyForm = Me.Cells(lCurRow, M_L_FORM_COL)
'             If "Long" = sPropertyForm Then
'                 vCompContent = 0
'             ElseIf "String" = sPropertyForm Then
'                 vCompContent = """"""
'             Else
'                 vCompContent = """"""
'             End If
'             sTmpCode = sTmpCode & Replace(M_S_VBA_INIT_PART_CONTENT, M_S_PROPERTY_NAME, sPropertyName)
'             sTmpCode = Replace(sTmpCode, M_S_COMP_CONTENT, vCompContent)
'         Next lCurRow
'         sTargetCode = M_S_VBA_CLASS_INIT_FORMAT
'         sTargetCode = Replace(sTargetCode, M_S_INIT_CONTENT, sTmpCode)
'         Print #1, sTargetCode
'     End If
    
'     'check format
'     If Me.Range(G_S_CELL_NAME_TARHET_PROGRAM_LANGUAGE) = G_S_VBA Then
'         sTmpCode = ""
'         For lCurRow = M_L_START_ROW To lEndRow
'             sPropertyName = Replace(Me.Cells(lCurRow, M_L_MEMBER_COL), "m_", "")
'             sPropertyForm = Me.Cells(lCurRow, M_L_FORM_COL)
'             If "Long" = sPropertyForm Then
'                 vCompContent = 0
'             ElseIf "String" = sPropertyForm Then
'                 vCompContent = """"""
'             Else
'                 vCompContent = """"""
'             End If
'             sTmpCode = sTmpCode & Replace(M_S_VBA_CHECK_PART_CONTENT, M_S_PROPERTY_NAME, sPropertyName)
'             sTmpCode = Replace(sTmpCode, M_S_COMP_CONTENT, vCompContent)
'         Next lCurRow
'         sTargetCode = M_S_VBA_PROPERTY_CHECK_FORMAT
'         sTargetCode = Replace(sTargetCode, M_S_CHECK_CONTENT, sTmpCode)
'         Print #1, sTargetCode
'     End If
    
'     Close #1
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnOpen
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnOpen_Click()
'     Dim oWscriptShell As Object
'     Set oWscriptShell = CreateObject("Wscript.Shell")
'     oWscriptShell.Run ThisWorkbook.Path & M_S_CODE_FILE_NAME, 3
' End Sub

' '--------------------------------------------------------------------------------
' '@brief : click btnClear
' '@param : NULL
' '@return: NULL
' '--------------------------------------------------------------------------------
' Public Sub btnClear_Click()
'     Call init

'     Dim RangeClearer As New cRangeClearer
'     With RangeClearer
'         Call .ClearRangeContents(m_siCreateProperty)
'     End With
' End Sub



















