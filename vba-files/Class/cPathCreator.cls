VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPathCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_bLock As Boolean
Private m_sFullPath As String
Private m_sFolderPath As String
Private m_sBaseName As String
Private m_sFileName As String
Private m_sFileExtension As String
Private m_FileFormatOperator As cFileFormatOperator

Public Property Get bLock() As Boolean
    bLock = m_bLock
End Property
Public Property Let bLock(ByVal vbLock As Boolean)
    m_bLock = vbLock
End Property

Public Property Get sFullPath() As String
    sFullPath = m_sFullPath
End Property
Public Property Let sFullPath(ByVal vsFullPath As String)
    m_sFullPath = vsFullPath
End Property

Public Property Get sFolderPath() As String
    sFolderPath = m_sFolderPath
End Property
Public Property Let sFolderPath(ByVal vsFolderPath As String)
    m_sFolderPath = vsFolderPath
End Property

Public Property Get sBaseName() As String
    sBaseName = m_sBaseName
End Property
Public Property Let sBaseName(ByVal vsBaseName As String)
    m_sBaseName = vsBaseName
End Property

Public Property Get sFileName() As String
    sFileName = m_sFileName
End Property
Public Property Let sFileName(ByVal vsFileName As String)
    m_sFileName = vsFileName
End Property

Public Property Get sFileExtension() As String
    sFileExtension = m_sFileExtension
End Property
Public Property Let sFileExtension(ByVal vsFileExtension As String)
    m_sFileExtension = vsFileExtension
End Property

' Initialize member variables.
Private Sub Class_Initialize()
    m_bLock = False
    m_sFullPath = ""
    m_sFolderPath = ""
    m_sBaseName = ""
    m_sFileName = ""
    m_sFileExtension = ""
    Set m_FileFormatOperator = New cFileFormatOperator
End Sub

Public Function SetPath() As Boolean
    SetPath = False
    If m_bLock Then: Exit Function

    Dim vPathArray As Variant
    Dim sFileName As String
    If "" <> m_sFullPath Then
        vPathArray = Split(m_sFullPath, "\")
        sFileName = vPathArray(UBound(vPathArray))
        If m_FileFormatOperator.HasExtension(sFileName) Then
            m_sFileName = sFileName
            m_sBaseName = m_FileFormatOperator.RemoveExtension(sFileName)
            m_sFileExtension = m_FileFormatOperator.GetFileFormatString(sFileName)
            SetPath = True
            Exit Function
        End If
    End If
    If "" <> m_sFolderPath And m_FileFormatOperator.HasExtension(m_sFileName) Then
        m_sFullPath = m_sFolderPath & "\" & m_sFileName
        m_sBaseName = m_FileFormatOperator.RemoveExtension(m_sFileName)
        m_sFileExtension = m_FileFormatOperator.GetFileFormatString(m_sFileName)
        SetPath = True
        Exit Function
    End If
    If "" <> m_sFolderPath And "" <> m_sBaseName _
    And m_FileFormatOperator.HasExtension(m_sFileExtension) Then
        m_sFullPath = m_sFolderPath & "\" & m_sBaseName & m_sFileExtension
        m_sFileName = m_sBaseName & m_sFileExtension
        SetPath = True
        Exit Function
    End If
End Function

' Receive enum of path number, then combine exisiting path and vsFileName value.
' If vsFileName is not exist, exit this function.
Public Function CombineFolderPath(ByVal vsFolderName As String) As Boolean
    CombineFolderPath = False
    If "" = vsFolderName Or m_bLock Then: Exit Function
    
    If "" <> m_sFolderPath Then
        m_sFolderPath = m_sFolderPath & "\" & vsFolderName
        CombineFolderPath = True
    Else
        m_sFolderPath = vsFolderName
        CombineFolderPath = True
    End If
End Function

