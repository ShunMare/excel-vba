VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSeparateString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_lSeparateType As Long
Dim m_lStartRow As Long
Dim m_lTargetCol As Long
Dim m_lInputCol As Long
Dim m_sSeparateKey As String
Dim m_sFrontSeparateKey As String
Dim m_sRearSeparateKey As String

'--------------------------------------------------------------------------------
'@brief : getter letter sWorkbookName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sWorkbookName() As String
    sWorkbookName = m_sWorkbookName
End Property

Public Property Let sWorkbookName(ByVal vsWorkbookName As String)
    m_sWorkbookName = vsWorkbookName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sWorksheetName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sWorksheetName() As String
    sWorksheetName = m_sWorksheetName
End Property

Public Property Let sWorksheetName(ByVal vsWorksheetName As String)
    m_sWorksheetName = vsWorksheetName
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lSeparateType
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lSeparateType() As Long
    lSeparateType = m_lSeparateType
End Property

Public Property Let lSeparateType(ByVal vlSeparateType As Long)
    m_lSeparateType = vlSeparateType
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lStartRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lStartRow() As Long
    lStartRow = m_lStartRow
End Property

Public Property Let lStartRow(ByVal vlStartRow As Long)
    m_lStartRow = vlStartRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetCol() As Long
    lTargetCol = m_lTargetCol
End Property

Public Property Let lTargetCol(ByVal vlTargetCol As Long)
    m_lTargetCol = vlTargetCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lInputCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lInputCol() As Long
    lInputCol = m_lInputCol
End Property

Public Property Let lInputCol(ByVal vlInputCol As Long)
    m_lInputCol = vlInputCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sSeparateKey
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sSeparateKey() As String
    sSeparateKey = m_sSeparateKey
End Property

Public Property Let sSeparateKey(ByVal vsSeparateKey As String)
    m_sSeparateKey = vsSeparateKey
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sFrontSeparateKey
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sFrontSeparateKey() As String
    sFrontSeparateKey = m_sFrontSeparateKey
End Property

Public Property Let sFrontSeparateKey(ByVal vsFrontSeparateKey As String)
    m_sFrontSeparateKey = vsFrontSeparateKey
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sRearSeparateKey
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sRearSeparateKey() As String
    sRearSeparateKey = m_sRearSeparateKey
End Property

Public Property Let sRearSeparateKey(ByVal vsRearSeparateKey As String)
    m_sRearSeparateKey = vsRearSeparateKey
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sWorkbookName = ""
    m_sWorksheetName = ""
    m_lSeparateType = 0
    m_lStartRow = 0
    m_lTargetCol = 0
    m_lInputCol = 0
    m_sSeparateKey = ""
    m_sFrontSeparateKey = ""
    m_sRearSeparateKey = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sWorksheetName Then
        checkMyValue = False
        Exit Function
    End If
    If E_SEPARATE_STRING.SEPARATE_STRING_NUM <= m_lSeparateType Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lStartRow Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lTargetCol Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lInputCol Then
        checkMyValue = False
        Exit Function
    End If
    If E_SEPARATE_STRING.NORMAL = m_lSeparateType Then
        If "" = m_sSeparateKey Then
            checkMyValue = False
            Exit Function
        End If
    End If
    If E_SEPARATE_STRING.BETWEEN = m_lSeparateType Then
        If "" = m_sFrontSeparateKey Then
            checkMyValue = False
            Exit Function
        End If
        If "" = m_sRearSeparateKey Then
            checkMyValue = False
            Exit Function
        End If
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : separate string
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub SeparateString()
    Dim wsTarget As Worksheet
    Dim lEndRow As Long
    
    If Not checkMyValue Then
        Exit Sub
    End If
    
    If sWorkbookName = "" Then
        Set wsTarget = ThisWorkbook.Worksheets(sWorksheetName)
    ElseIf wsTarget <> "" Then
        Set wsTarget = Workbooks(sWorkbookName).Worksheets(sWorksheetName)
    Else
        MsgBox "err"
        Exit Sub
    End If
    
    lEndRow = wsTarget.Cells(wsTarget.Rows.Count, lTargetCol).End(xlUp).Row
    If E_SEPARATE_STRING.NORMAL = m_lSeparateType Then
        Call SeparateStringNormal(wsTarget, lEndRow)
    ElseIf E_SEPARATE_STRING.BETWEEN = m_lSeparateType Then
        Call SeparateStringBetween(wsTarget, lEndRow)
    End If
End Sub

'--------------------------------------------------------------------------------
'@brief : separate string normal
'@param : vwsTarget
'@param : vlEndRow
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub SeparateStringNormal(ByVal vwsTarget As Worksheet, ByVal vlEndRow As Long)
    Dim lCurRow As Long
    Dim sOriginal As String
    Dim vArray As Variant
    Dim lCurRow As Long
    Dim lCntArray As Long

    For lCurRow = lStartRow To vlEndRow
        vArray = Split(Cells(lCurRow, lTargetCol), sSeparateKey)
        For lCntArray = LBound(vArray) To UBound(vArray)
            Cells(lCurRow, lInputCol + lCntArray + 1) = vArray(lCntArray)
        Next
    Next lCurRow
End Sub

'--------------------------------------------------------------------------------
'@brief : separate string between
'@param : vwsTarget
'@param : vlEndRow
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub SeparateStringBetween(ByVal vwsTarget As Worksheet, ByVal vlEndRow As Long)
    Dim lCurRow As Long
    Dim lFrontPso As Long
    Dim lRearPso As Long
    Dim lTargetStrLen As Long
    Dim sOriginal As String
    
    For lCurRow = lStartRow To vlEndRow
        sOriginal = vwsTarget.Cells(lCurRow, lTargetCol)
        lFrontPso = InStr(sOriginal, sFrontSeparateKey) + Len(sFrontSeparateKey)
        lRearPso = InStr(sOriginal, sRearSeparateKey) - 1
        lTargetStrLen = lRearPso - (lFrontPso - 1)
        vwsTarget.Cells(lCurRow, lInputCol) = Mid(sOriginal, lFrontPso, lTargetStrLen)
    Next lCurRow
End Sub
