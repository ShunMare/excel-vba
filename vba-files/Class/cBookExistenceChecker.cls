VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBookExistenceChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_FileFormatOperator As cFileFormatOperator

' Initialize member variables.
Private Sub Class_Initialize()
    Set m_FileFormatOperator = New cFileFormatOperator
End Sub

' Check member variables.
' Judge whether to proceed to the next process.
Private Function checkMyValue()
    checkMyValue = True
End Function

' Check whether workbook is exist in opening workbooks.
' Before call this function, should set bellow member variables :
' m_pcTarget.sFileName        : Target file name.
' If target book name is not exist, return false.
Public Function IsBookInOpen(ByVal vpcTarget As cPathCreator) As Boolean
    Dim sFileName As String
    Dim wbTarget As Workbook
    For Each wbTarget In Workbooks
        If wbTarget.Name = vpcTarget.sFileName Then
            IsBookInOpen = True
            Exit Function
        End If
    Next wbTarget
    IsBookInOpen = False
End Function
