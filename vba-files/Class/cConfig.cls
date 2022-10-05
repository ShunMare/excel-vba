VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_sDROExportPath As String
Dim m_sLocalDrive As String

'--------------------------------------------------------------------------------
'@brief : getter letter sDROExportPath
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sDROExportPath() As String
    sDROExportPath = m_sDROExportPath
End Property

Public Property Let sDROExportPath(ByVal vsDROExportPath As String)
    m_sDROExportPath = vsDROExportPath
End Property

'--------------------------------------------------------------------------------
'@brief : getter letter sLocalDrive
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get sLocalDrive() As String
    sLocalDrive = m_sLocalDrive
End Property

Public Property Let sLocalDrive(ByVal vsLocalDrive As String)
    m_sLocalDrive = vsLocalDrive
End Property
