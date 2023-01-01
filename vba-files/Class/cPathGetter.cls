VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPathGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function checkMyValue()
    checkMyValue = True
End Function

Public Function GetFolderPathDialog() As String
    Dim oFileDialog As Object
    Dim sFolderPath As String
    Set oFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If oFileDialog.Show Then
        GetFolderPathDialog = oFileDialog.SelectedItems(1)
    End If
End Function

Public Function GetDrive() As String()
    Dim lCntAsc As Long
    Dim sDriveArray() As String
    Dim lCntArray As Long: lCntArray = 0
    For lCntAsc = Asc("A") To Asc("Z")
        If "" <> Dir(Chr(lCntAsc) & ":\", vbDirectory) Then
            ReDim Preserve sDriveArray(lCntArray)
            sDriveArray(lCntArray) = Chr(lCntAsc) & ":"
            lCntArray = lCntArray + 1
        End If
    Next lCntAsc
    GetDrive = sDriveArray
End Function

