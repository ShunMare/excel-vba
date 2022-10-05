VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMessageShower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sReplaceWord1 As String
Dim m_sReplaceWord2 As String

'--------------------------------------------------------------------------------
'@brief : getter letter sReplaceWord1
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sReplaceWord1() As String
    sReplaceWord1 = m_sReplaceWord1
End Property

Public Property Let sReplaceWord1(ByVal vsReplaceWord1 As String)
    m_sReplaceWord1 = vsReplaceWord1
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sReplaceWord2
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sReplaceWord2() As String
    sReplaceWord2 = m_sReplaceWord2
End Property

Public Property Let sReplaceWord2(ByVal vsReplaceWord2 As String)
    m_sReplaceWord2 = vsReplaceWord2
End Property

'--------------------------------------------------------------------------------
'@brief : peplace word
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Function replaceWord(ByVal sMsg As String)
    If InStr(sMsg, G_S_MSG_REPLACE_1) <> 0 Then
        sMsg = Replace(sMsg, G_S_MSG_REPLACE_1, sReplaceWord1)
    End If
    If InStr(sMsg, G_S_MSG_REPLACE_2) <> 0 Then
        sMsg = Replace(sMsg, G_S_MSG_REPLACE_2, sReplaceWord2)
    End If
    replaceWord = sMsg
End Function

'--------------------------------------------------------------------------------
'@brief : clean up
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub CleanUpWords()
    sReplaceWord1 = ""
    sReplaceWord2 = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : show information message
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ShowExclamationMsg(ByVal vsMsgNo As String)
    MsgBox replaceWord(vsMsgNo), vbExclamation
End Sub

'--------------------------------------------------------------------------------
'@brief : show information message
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ShowInformationMsg(ByVal vsMsgNo As String)
    MsgBox replaceWord(vsMsgNo), vbInformation
End Sub

'--------------------------------------------------------------------------------
'@brief : show critical message
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ShowCriticalMsg(ByVal vsMsgNo As String)
    MsgBox replaceWord(vsMsgNo), vbCritical
End Sub
