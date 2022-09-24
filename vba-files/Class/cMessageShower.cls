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

'--------------------------------------------------------------------------------
'@brief : show error message
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ShowErrorMessage(ByVal sErrorMsgNo As String)
    MsgBox sErrorMsgNo, vbCritical
End Sub
