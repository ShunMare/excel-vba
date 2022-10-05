VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cValueInRangeGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lItemStartrow As Long
Dim m_lItemStartCol As Long
Dim m_lItemEndRow As Long
Dim m_lItemEndCol As Long
Dim m_lTargetRow As Long
Dim m_lTargetCol As Long
Dim m_rTarget As Range
Dim m_sTargetItem As String

'--------------------------------------------------------------------------------
'@brief : getter letter lItemStartRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lItemStartRow() As Long
    Let lItemStartRow = m_lItemStartrow
End Property

Public Property Let lItemStartRow(ByVal vlItemStartRow As Long)
    Let m_lItemStartrow = vlItemStartRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lItemStartCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lItemStartCol() As Long
    Let lItemStartCol = m_lItemStartCol
End Property

Public Property Let lItemStartCol(ByVal vlItemStartCol As Long)
    Let m_lItemStartCol = vlItemStartCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lItemEndRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lItemEndRow() As Long
    Let lItemEndRow = m_lItemEndRow
End Property

Public Property Let lItemEndRow(ByVal vlItemEndRow As Long)
    Let m_lItemEndRow = vlItemEndRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lItemEndCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lItemEndCol() As Long
    Let lItemEndCol = m_lItemEndCol
End Property

Public Property Let lItemEndCol(ByVal vlItemEndCol As Long)
    Let m_lItemEndCol = vlItemEndCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetRow() As Long
    Let lTargetRow = m_lTargetRow
End Property

Public Property Let lTargetRow(ByVal vlTargetRow As Long)
    Let m_lTargetRow = vlTargetRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetCol() As Long
    Let lTargetCol = m_lTargetCol
End Property

Public Property Let lTargetCol(ByVal vlTargetCol As Long)
    Let m_lTargetCol = vlTargetCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter setter rTarget
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get rTarget() As Range
    Set rTarget = m_rTarget
End Property

Public Property Set rTarget(ByVal vrTarget As Range)
    Set m_rTarget = vrTarget
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sTargetItem
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetItem() As String
    sTargetItem = m_sTargetItem
End Property

Public Property Let sTargetItem(ByVal vsTargetItem As String)
    m_sTargetItem = vsTargetItem
End Property

'--------------------------------------------------------------------------------
'@brief : set my value
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function SetMyValue() As Long
    lItemEndRow = IIf(lItemEndRow = 0, rTarget.Rows.Count, lTargetCol)
End Function

'--------------------------------------------------------------------------------
'@brief : get target value in range
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function GetValueInRange() As Variant
    Dim lCurRow As Long
    Dim lCurItemCol As Long
    Dim lCurTargetCol As Long
    Dim sMyItem As String
    Dim sCellAddress As String
    Dim wsTarget As Worksheet

    sMyItem = ""
    Set wsTarget = rTarget.Parent
    Dim MergeAreaValueGetter As New cMergeAreaValueGetter
    With MergeAreaValueGetter
        .lConnectionProcessing = E_CP.COMBINING
        For lCurRow = lItemStartRow To lItemEndRow
            For lCurItemCol = lItemStartCol To lItemEndCol
                .lRetType = vbString
                sCellAddress = rTarget.Cells(lCurRow, lCurItemCol).Address
                sMyItem = sMyItem & .GetMergeAreaValue(wsTarget.Range(sCellAddress))
            Next lCurItemCol
            If sTargetItem = sMyItem Then
                .lRetType = vbDouble
                sCellAddress = rTarget.Cells(lCurRow, lTargetCol).Address
                GetValueInRange = .GetMergeAreaValue(wsTarget.Range(sCellAddress))
                Exit Function
            Else
                sMyItem = ""
            End If
        Next lCurRow
    End With
End Function

'--------------------------------------------------------------------------------
'@brief : get item in range
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function GetItemInRange() As String
    Dim lCurCol As Long
    Dim sItemInRange As String
    Dim sCellAddress As String
    Dim wsTarget As Worksheet
    
    Set wsTarget = rTarget.Parent
    sItemInRange = ""
    Dim MergeAreaValueGetter As New cMergeAreaValueGetter
    With MergeAreaValueGetter
        .lRetType = vbString
        .lConnectionProcessing = E_CP.COMBINING
        For lCurCol = lItemStartCol To lItemEndCol
            sCellAddress = rTarget.Cells(lTargetRow, lCurCol).Address
            sItemInRange = sItemInRange & .GetMergeAreaValue(wsTarget.Range(sCellAddress))
        Next lCurCol
    End With
    GetItemInRange = sItemInRange
End Function
