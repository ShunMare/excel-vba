VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExportVBAFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ComponentType
    STANDARD_MODULE = 1
    CLASS_MODULE = 2
    USER_FORM = 3
End Enum

'--------------------------------------------------------------------------------
'@brief : export VBA files
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub ExportVBAFiles()
    Dim oVBAComponent As Object
    Dim sExportPath As String
    
    sExportPath = ThisWorkbook.Path & "\bin"
    If Dir(sExportPath, vbDirectory) = "" Then
        Call MkDir(sExportPath)
    End If

    If Dir(sExportPath & "\bas", vbDirectory) = "" Then
        Call MkDir(sExportPath & "\bas")
    End If
    If Dir(sExportPath & "\cls", vbDirectory) = "" Then
        Call MkDir(sExportPath & "\cls")
    End If
    If Dir(sExportPath & "\frm", vbDirectory) = "" Then
        Call MkDir(sExportPath & "\frm")
    End If

    For Each oVBAComponent In ThisWorkbook.VBProject.VBComponents
        Select Case oVBAComponent.Type
        Case STANDARD_MODULE
            oVBAComponent.Export sExportPath & "\bas\" & oVBAComponent.Name & ".bas"
        Case CLASS_MODULE
            oVBAComponent.Export sExportPath & "\cls\" & oVBAComponent.Name & ".cls"
        Case USER_FORM
            oVBAComponent.Export sExportPath & "\frm\" & oVBAComponent.Name & ".frm"
        End Select
    Next
End Sub
