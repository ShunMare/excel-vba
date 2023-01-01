VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileNameOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Initialize member variables.
Private Sub Class_Initialize()
End Sub

' Check member variables.
' Judge whether to proceed to the next process.
Private Function checkMyValue()
    checkMyValue = True
End Function

'
Public Function GetFileCount(ByVal vpcTarget As cPathCreator, _
 Optional ByVal sNotIncludedFileName As String = "") As Long
    Dim sFileName As String
    sFileName = Dir(vpcTarget.sFolderPath & "\*" & vpcTarget.sFileExtension)
    
    Dim lCntFile As Long: lCntFile = 0
    Dim bTarget As Boolean: bTarget = False
    Do While sFileName <> ""
        If "" = sNotIncludedFileName Then
            bTarget = True
        ElseIf Not sFileName Like "*" & sNotIncludedFileName & "*" Then
            bTarget = True
        End If
        If bTarget Then
            lCntFile = lCntFile + 1
            bTarget = False
        End If
        sFileName = Dir()
    Loop
    GetFileCount = lCntFile
End Function


'
Public Function GetFileNameArray(ByVal vpcTarget As cPathCreator, _
Optional ByVal sNotIncludedFileName As String = "") As String()
    Dim sFileName As String
    sFileName = Dir(vpcTarget.sFolderPath & "\*" & vpcTarget.sFileExtension)
    
    Dim sFileNameArray() As String
    Dim lCntArray As Long: lCntArray = 0
    Dim bTarget As Boolean: bTarget = False
    Do While sFileName <> ""
        If "" = sNotIncludedFileName Then
            bTarget = True
        ElseIf Not sFileName Like "*" & sNotIncludedFileName & "*" Then
            bTarget = True
        End If
        If bTarget Then
            ReDim Preserve sFileNameArray(lCntArray)
            sFileNameArray(lCntArray) = sFileName
            lCntArray = lCntArray + 1
            bTarget = False
        End If
        sFileName = Dir()
    Loop
    GetFileNameArray = sFileNameArray
End Function

'
Public Function GetFileName(ByVal vpcTarget As cPathCreator, _
ByVal vlIndexNo As Long, Optional ByVal sNotIncludedFileName As String = "") As String
    Dim sFileName As String
    sFileName = Dir(vpcTarget.sFolderPath & "\*" & vpcTarget.sFileExtension)
    
    Dim lCntIndexNo As Long: lCntIndexNo = 0
    Dim bTarget As Boolean: bTarget = False
    Do While sFileName <> ""
        If "" = sNotIncludedFileName Then
            bTarget = True
        ElseIf Not sFileName Like "*" & sNotIncludedFileName & "*" Then
            bTarget = True
        End If
        If bTarget Then
            If vlIndexNo = lCntIndexNo Then
                GetFileName = sFileName
                Exit Function
            Else
                lCntIndexNo = lCntIndexNo + 1
                bTarget = False
            End If
        End If
        sFileName = Dir()
    Loop
End Function

'
Public Function ChangeFileNameArray(ByVal vpcTarget As cPathCreator, _
vsPreFileName() As String, vsPostFileName() As String) As Boolean
    ChangeFileNameArray = False
    If UBound(vsPreFileName) <> UBound(vsPostFileName) Then: Exit Function
    
    Dim lCntArray As Long
    For lCntArray = LBound(vsPreFileName) To UBound(vsPreFileName) - 1
        vpcTarget.sFileName = vsPreFileName(lCntArray)
        vpcTarget.sFullPath = ""
        Call vpcTarget.SetPath
        If Not ChangeFileName(vpcTarget, vsPostFileName(lCntArray)) Then
            Exit Function
        End If
    Next lCntArray
    ChangeFileNameArray = True
End Function

'
Public Function ChangeFileName(ByVal vpcTarget As cPathCreator, _
ByVal vsPostFileName As String) As Boolean
    ChangeFileName = False
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    Dim oFileSystemObject As Object
    Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
    If "" <> Dir(vpcTarget.sFolderPath & vsPostFileName) Then: Exit Function
    oFileSystemObject.GetFile(vpcTarget.sFullPath).Name = vsPostFileName
    ChangeFileName = True
End Function

