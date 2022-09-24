VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFolderPathInputter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sCellName As String

'--------------------------------------------------------------------------------
'@brief : getter letter sCellName
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sCellName() As String
    sCellName = m_sCellName
End Property

Public Property Let sCellName(ByVal vsCellName As String)
    m_sCellName = vsCellName
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sCellName = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sCellName Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : get folder path
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function GetFolderPath()
    Dim objFileDialog As Object
    Dim sFolderPath As String
    
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    If objFileDialog.Show = True Then
        sFolderPath = objFileDialog.SelectedItems(1)
    Else
        ''do nothing
    End If
    Set objFileDialog = Nothing
    GetFolderPath = sFolderPath
End Function

'--------------------------------------------------------------------------------
'@brief : input folder path
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub InputFolderPath(siTarget As cSheetInfo)
    Dim rTarget As Range
    
    If Not checkMyValue Then
        Exit Sub
    End If
    
    Set rTarget = siTarget.ws.Range(sCellName)
    rTarget = GetFolderPath
End Sub

