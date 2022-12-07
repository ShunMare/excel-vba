VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileNameChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sTargetFolderPath As String
Dim m_lChangeFileNameCol As Long
Dim m_lCurFileNameCol As Long
Dim m_MessageShower As cMessageShower

'--------------------------------------------------------------------------------
'@brief : getter letter sTargetFolderPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sTargetFolderPath() As String
    sTargetFolderPath = m_sTargetFolderPath
End Property

Public Property Let sTargetFolderPath(ByVal vsTargetFolderPath As String)
    m_sTargetFolderPath = vsTargetFolderPath
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lChangeFileNameCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lChangeFileNameCol() As Long
    lChangeFileNameCol = m_lChangeFileNameCol
End Property

Public Property Let lChangeFileNameCol(ByVal vlChangeFileNameCol As Long)
    m_lChangeFileNameCol = vlChangeFileNameCol
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter lCurFileNameCol
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lCurFileNameCol() As Long
    lCurFileNameCol = m_lCurFileNameCol
End Property

Public Property Let lCurFileNameCol(ByVal vlCurFileNameCol As Long)
    m_lCurFileNameCol = vlCurFileNameCol
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_sTargetFolderPath = ""
    m_lChangeFileNameCol = 0
    m_lCurFileNameCol = 0
End Sub

'--------------------------------------------------------------------------------
'@brief : check my value
'@param : NULL
'@return: True, False
'--------------------------------------------------------------------------------
Private Function checkMyValue()
    If "" = m_sTargetFolderPath Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lChangeFileNameCol Then
        checkMyValue = False
        Exit Function
    End If
    If 0 = m_lCurFileNameCol Then
        checkMyValue = False
        Exit Function
    End If
    checkMyValue = True
End Function

'--------------------------------------------------------------------------------
'@brief : change file name
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ChangeFileName(ByVal siTarget As cSheetInfo)
    Dim lCurRow As Long
    Dim sTargetFilePath As String
    Dim sTargetFileName As String
    Dim sChangeFileName As String
    Dim oFileSystemObject As Object
    
    If Not checkMyValue Then
        m_MessageShower.ShowErrorMessage (G_S_ERROR0000)
        Exit Sub
    End If
    
    Dim ResultInputter As New cResultInputter
    Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    For lCurRow = siTarget.lStartRow To siTarget.lEndRow
        sTargetFileName = siTarget.ws.Cells(lCurRow, lCurFileNameCol)
        sTargetFilePath = sTargetFolderPath & "\" & sTargetFileName
        sChangeFileName = siTarget.ws.Cells(lCurRow, lChangeFileNameCol)
        
        With ResultInputter
            .lTargetRow = lCurRow
            If Not sChangeFileName Like "*.*" Then
                .sMsg = "not extension"
                Call .InputResult(siTarget)
            ElseIf Not oFileSystemObject.FileExists(sTargetFilePath) Then
                .sMsg = "not exit"
                Call .InputResult(siTarget)
            ElseIf sTargetFileName = sChangeFileName Then
                .sMsg = "not necessary"
                Call .InputResult(siTarget)
            Else
                oFileSystemObject.GetFile(sTargetFilePath).Name = sChangeFileName
                .sMsg = ""
                Call .InputResult(siTarget)
            End If
        End With
    Next lCurRow
End Sub

