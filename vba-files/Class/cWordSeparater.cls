VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWordSeparater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lSeparateType As Long
Dim m_sSeparateKey As String
Dim m_sFrontSeparateKey As String
Dim m_sRearSeparateKey As String

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
    m_lSeparateType = 0
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
    If E_SEPARATE_STRING.SEPARATE_STRING_NUM <= m_lSeparateType Then
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
Public Sub SeparateString(ByVal siTarget As cSheetInfo)
    If Not checkMyValue Then
        Exit Sub
    End If
    
    If E_SEPARATE_STRING.NORMAL = m_lSeparateType Then
        Call SeparateStringNormal(siTarget)
    ElseIf E_SEPARATE_STRING.BETWEEN = m_lSeparateType Then
        Call SeparateStringBetween(siTarget)
    End If
End Sub

'--------------------------------------------------------------------------------
'@brief : separate string normal
'@param : vwsTarget
'@param : vlEndRow
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub SeparateStringNormal(ByVal siTarget As cSheetInfo)
    Dim lCurRow As Long
    Dim sOriginal As String
    Dim vArray As Variant
    Dim lCntArray As Long

    For lCurRow = siTarget.lStartRow To siTarget.lEndRow
        vArray = Split(siTarget.ws.Cells(lCurRow, siTarget.lTargetCol), sSeparateKey)
        For lCntArray = LBound(vArray) To UBound(vArray)
            siTarget.ws.Cells(lCurRow, siTarget.lTargetCol + lCntArray) = Trim(vArray(lCntArray))
        Next
    Next lCurRow
End Sub

'--------------------------------------------------------------------------------
'@brief : separate string between
'@param : vwsTarget
'@param : vlEndRow
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub SeparateStringBetween(ByVal siTarget As cSheetInfo)
    Dim lCurRow As Long
    Dim lFrontPso As Long
    Dim lRearPso As Long
    Dim lTargetStrLen As Long
    Dim sOriginal As String

    If Not checkMyValue Then
        Exit Sub
    End If

    For lCurRow = siTarget.lStartRow To siTarget.lEndRow
        sOriginal = siTarget.ws.Cells(lCurRow, siTarget.lTargetCol)
        lFrontPso = InStr(sOriginal, sFrontSeparateKey) + Len(sFrontSeparateKey)
        lRearPso = InStr(sOriginal, sRearSeparateKey) - 1
        lTargetStrLen = lRearPso - (lFrontPso - 1)
        siTarget.ws.Cells(lCurRow, siTarget.lTargetCol) = Mid(sOriginal, lFrontPso, lTargetStrLen)
    Next lCurRow
End Sub
