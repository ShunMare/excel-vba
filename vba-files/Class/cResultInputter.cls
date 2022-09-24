VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cResultInputter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lTargetRow As Long
Dim m_sMsg As String

'--------------------------------------------------------------------------------
'@brief : getter letter lTargetRow
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lTargetRow() As Long
    lTargetRow = m_lTargetRow
End Property

Public Property Let lTargetRow(ByVal vlTargetRow As Long)
    m_lTargetRow = vlTargetRow
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sMsg
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sMsg() As String
    sMsg = m_sMsg
End Property

Public Property Let sMsg(ByVal vsMsg As String)
    m_sMsg = vsMsg
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_lTargetRow = 0
    m_sMsg = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : input result
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub InputResult(ByVal siTarget As cSheetInfo)
    Dim rTarget As Range
    
    If Not checkMyValue Then
        Exit Sub
    End If

    With siTarget
        Set rTarget = .ws.Cells(lTargetRow, .lResultCol)
    End With
    
    If sMsg <> "" Then
        rTarget = sMsg
    Else
        rTarget = "TRUE"
    End If
End Sub
