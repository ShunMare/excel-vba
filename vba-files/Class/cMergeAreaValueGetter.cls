VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMergeAreaValueGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''ConnectionProcessing = CP
Public Enum E_CP
    OVER_WRITE
    ADDITION
    SUBTRACTION
    MULTIPICATION
    DIVITION
    COMBINING
End Enum

Dim m_lRetType As Long
Dim m_lConnectionProcessing As Long
Dim m_vRetValue As Variant

'--------------------------------------------------------------------------------
'@brief : getter letter lRetType
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lRetType() As Long
    Let lRetType = m_lRetType
End Property

Public Property Let lRetType(ByVal vlRetType As Long)
    Let m_lRetType = vlRetType
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lConnectionProcessing
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lConnectionProcessing() As Long
    Let lConnectionProcessing = m_lConnectionProcessing
End Property

Public Property Let lConnectionProcessing(ByVal vlConnectionProcessing As Long)
    Let m_lConnectionProcessing = vlConnectionProcessing
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter vRetValue
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get vRetValue() As Variant
    Let vRetValue = m_vRetValue
End Property

Public Property Let vRetValue(ByVal vvRetValue As Variant)
    Let m_vRetValue = vvRetValue
End Property

'--------------------------------------------------------------------------------
'@brief : get merge area value
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function GetMergeAreaValue(ByVal rTarget As Range) As String
    Dim lCntCellItem As Long
    Dim lCurRow As Long
    Dim lStartRow As Long
    Dim lEndRow As Long
    Dim lCurCol As Long
    Dim lStartCol As Long
    Dim lEndCol As Long
    
    lCntCellItem = 0
    vRetValue = ""
    With rTarget.MergeArea
        lStartRow = rTarget.MergeArea.Item(1).Row
        lEndRow = rTarget.MergeArea.Item(.Count).Row
        lStartCol = rTarget.MergeArea.Item(1).Column
        lEndCol = rTarget.MergeArea.Item(.Count).Column
        For lCurRow = lStartRow To lEndRow
            For lCurCol = lStartCol To lEndCol
                lCntCellItem = lCntCellItem + 1
                Call ProcessValue(.Item(lCntCellItem))
            Next lCurCol
        Next lCurRow
    End With
    GetMergeAreaValue = vRetValue
End Function

'--------------------------------------------------------------------------------
'@brief : process value
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ProcessValue(ByVal vValue As Variant)
    
    Select Case lConnectionProcessing
    Case E_CP.ADDITION
        vRetValue = vRetValue + vValue
    Case E_CP.SUBTRACTION
        vRetValue = vRetValue - vValue
    Case E_CP.MULTIPICATION
        vRetValue = vRetValue * vValue
    Case E_CP.DIVITION
        vRetValue = vRetValue / vValue
    Case E_CP.OVER_WRITE
        If Trim(vValue) <> "" Or vValue <> 0 Then
            vRetValue = vValue
        End If
    Case E_CP.COMBINING
        vRetValue = vRetValue & vValue
    End Select
End Sub

