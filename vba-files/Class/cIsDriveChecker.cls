VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIsDriveChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_lDrive  As Long
Dim m_sDrive  As String

'--------------------------------------------------------------------------------
'@brief : getter letter lDrive
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get lDrive() As String
    lDrive = m_lDrive
End Property

Public Property Let lDrive(ByVal vlDrive As String)
    m_lDrive = vlDrive
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sDrive
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sDrive() As String
    sDrive = m_sDrive
End Property

Public Property Let sDrive(ByVal vsDrive As String)
    m_sDrive = vsDrive
End Property

'--------------------------------------------------------------------------------
'@brief : initialize
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Private Sub Class_initialize()
    m_lDrive = 0
    m_sDrive = ""
End Sub

'--------------------------------------------------------------------------------
'@brief : check is drive
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub checkIsDrive()
    Dim lCntAsc As Long
    
    For lCntAsc = Asc("A") To Asc("Z")
        If Dir(Chr(lCntAsc) & ":\", vbDirectory) <> "" Then
            m_lDrive = lCntAsc
            m_sDrive = Chr(lCntAsc) & ":"
            Exit Sub
        End If
    Next lCntAsc
End Sub
