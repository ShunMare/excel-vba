VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWordSeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sSearchWord As String
Dim m_sTargetWord As String
Dim m_bExist As Boolean

'--------------------------------------------------------------------------------
'@brief : getter letter sSearchWord
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sSearchWord() As String
    sSearchWord = m_sSearchWord
End Property

Public Property Let sSearchWord(ByVal vsSearchWord As String)
    m_sSearchWord = vsSearchWord
End Property

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
'@brief : getter letter bExist
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get bExist() As Boolean
    bExist = m_bExist
End Property

Public Property Let bExist(ByVal vbExist As Boolean)
    m_bExist = vbExist
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sSearchWord = ""
    m_sTargetWord = ""
    m_bExist = False
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
'@brief : GetIsSearchWord
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub GetIsSearchWord()
    If Not checkMyValue Then
        Exit Sub
    End If
        
    If 0 <> InStr(sTargetWord, sSearchWord) And sSearchWord <> "" Then
        bExist = True
    Else
        bExist = False
    End If
End Sub
