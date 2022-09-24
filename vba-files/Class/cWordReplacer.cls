VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWordReplacer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sTargetWord As String
Dim m_sReplaceWord As String

'--------------------------------------------------------------------------------
'@brief : getter letter sTargetWord
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetWord() As String
    sTargetWord = m_sTargetWord
End Property

Public Property Let sTargetWord(ByVal vsTargetWord As String)
    m_sTargetWord = vsTargetWord
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sReplaceWord
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sReplaceWord() As String
    sReplaceWord = m_sReplaceWord
End Property

Public Property Let sReplaceWord(ByVal vsReplaceWord As String)
    m_sReplaceWord = vsReplaceWord
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sTargetWord = ""
    m_sReplaceWord = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sTargetWord Then
        checkMyValue = False
        Exit Function
    End If
    If "" = m_sReplaceWord Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : replace words
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub replaceWords(ByVal siTarget As cSheetInfo)
    Dim rTarget As Range
    
    If Not checkMyValue Then
        Exit Sub
    End If
    
    With siTarget
        Set rTarget = .ws.Range(.ws.Cells(.lStartRow, .lStartCol), .ws.Cells(.lEndRow, .lEndCol))
        rTarget.Replace What:=sTargetWord, Replacement:=sReplaceWord, LookAt:=xlWhole, MatchCase:=True
    End With
End Sub
