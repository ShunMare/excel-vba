VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRangeClearer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_MessageShower As cMessageShower

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    Set m_MessageShower = New cMessageShower
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
'@brief : clear range contents
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ClearRangeContents(siTarget As cSheetInfo)
    Dim rTarget As Range

    If Not checkMyValue Then
        m_MessageShower.ShowCriticalMsg (G_S_MSG0000)
        Exit Sub
    End If
    
    With siTarget
        Set rTarget = .ws.Range(.ws.Cells(.lStartRow, .lStartCol), .ws.Cells(.lEndRow, .lEndCol))
        rTarget.ClearContents
    End With
End Sub
