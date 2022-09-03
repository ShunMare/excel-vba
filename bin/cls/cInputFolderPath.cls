VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInputFolderPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sWorkbookName As String
Dim m_sWorksheetName As String
Dim m_sCellName As String

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get sWorkbookName() As String
    sWorkbookName = m_sWorkbookName
End Property

Public Property Let sWorkbookName(ByVal vsWorkbookName As String)
    m_sWorkbookName = vsWorkbookName
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get sWorksheetName() As String
    sWorksheetName = m_sWorksheetName
End Property

Public Property Let sWorksheetName(ByVal vsWorksheetName As String)
    m_sWorksheetName = vsWorksheetName
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Property Get sCellName() As String
    sCellName = m_sCellName
End Property

Public Property Let sCellName(ByVal vsCellName As String)
    m_sCellName = vsCellName
End Property

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Call Init
End Sub

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Sub Init()
    m_sWorkbookName = ""
    m_sWorksheetName = ""
    m_sCellName = ""
End Sub

'--------------------------------------------------------------------------------
'@brief :
'@param :
'@return:
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
'@brief :
'@param :
'@return:
'--------------------------------------------------------------------------------
Public Sub InputFolderPath()
    Dim rTarget As Range
    
    If sWorkbookName = "" Then
        Set rTarget = ThisWorkbook.Worksheets(sWorksheetName).Range(sCellName)
    ElseIf sWorkbookName <> "" Then
        Set rTarget = Workbooks(sWorkbookName).Worksheets(sWorksheetName).Range(sCellName)
    Else
        MsgBox "err"
        Exit Sub
    End If
    rTarget = GetFolderPath
End Sub

