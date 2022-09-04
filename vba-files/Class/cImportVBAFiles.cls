VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImportVBAFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ImportVBAFiles()
    Dim oVBAComponent As Object
    Dim sImportPath As String
    Dim oFileSystemObject As Object
    Dim oGetFiles As Object
    Dim oTargetFile As Object
    Dim sModukeName As String
    
    sImportPath = ThisWorkbook.Path & "\bin"
    If Dir(sImportPath, vbDirectory) = "" Then
        Call MkDir(sImportPath)
    End If
    
    If Dir(sImportPath & "\bas", vbDirectory) = "" Then
        Call MkDir(sImportPath & "\bas")
    End If
    If Dir(sImportPath & "\cls", vbDirectory) = "" Then
        Call MkDir(sImportPath & "\cls")
    End If
    If Dir(sImportPath & "\frm", vbDirectory) = "" Then
        Call MkDir(sImportPath & "\frm")
    End If
 
    Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
    Set oTargetFile = CreateObject("Scripting.FileSystemObject")
    
    Set oGetFiles = oFileSystemObject.GetFolder(sImportPath & "\bas")
    For Each oTargetFile In oGetFiles.Files
        sModukeName = Replace(Dir(oTargetFile), ".bas", "")
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(sModukeName)
        Error.Clear
        ThisWorkbook.VBProject.VBComponents.Import (oTargetFile)
    Next
    Set oGetFiles = oFileSystemObject.GetFolder(sImportPath & "\cls")
    For Each oTargetFile In oGetFiles.Files
        sModukeName = Replace(Dir(oTargetFile), ".cls", "")
        If sModukeName <> "cExportVBAFiles" And sModukeName <> "cImportVBAFiles" Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(sModukeName)
            Error.Clear
            ThisWorkbook.VBProject.VBComponents.Import (oTargetFile)
        End If
    Next
    Set oGetFiles = oFileSystemObject.GetFolder(sImportPath & "\frm")
    For Each oTargetFile In oGetFiles.Files
        sModukeName = Replace(Dir(oTargetFile), ".frm", "")
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(sModukeName)
        Error.Clear
        ThisWorkbook.VBProject.VBComponents.Import (oTargetFile)
    Next
End Sub

