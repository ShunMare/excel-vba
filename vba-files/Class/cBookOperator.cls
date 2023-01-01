VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBookOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is to operate Excel book.
'
'
' When use this class, should import bellow class:
' - cPathCreator
' - cBookExistenceChecker

Private m_BookExistenceChecker As cBookExistenceChecker

' Initialize member variables.
Private Sub Class_Initialize()
    Set m_BookExistenceChecker = New cBookExistenceChecker
End Sub

' Check member variables.
' Judge whether to proceed to the next process.
Private Function checkMyValue()
    checkMyValue = True
End Function

' Create book.
' Before call this function, should set bellow member variables :
' vpcTarget.sFullPath   : Created file is generated in the specified path.
' If specified folder path is not exist, this process finish.
Public Function CreateBook(ByVal vpcTarget As cPathCreator) As Boolean
    If "" <> Dir(vpcTarget.sFullPath) Then
        CreateBook = False
        Exit Function
    End If
    
    Dim wbCreate As Workbook
    Set wbCreate = Workbooks.Add
    Call wbCreate.SaveAs(vpcTarget.sFullPath)
    Application.DisplayAlerts = False
    wbCreate.Close
    Application.DisplayAlerts = True
    CreateBook = True
End Function

' Open book.
' Before call this function, should set bellow member variables :
' vpcTarget.sFullPath   : Book of specified path is opened.
' If specified folder path is not exist, this process finish.
' Book is opened return True.
Public Function OpenBook(ByVal vpcTarget As cPathCreator) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    If m_BookExistenceChecker.IsBookInOpen(vpcTarget) Then
        OpenBook = False
        Exit Function
    End If
    Workbooks.Open (vpcTarget.sFullPath)
    OpenBook = True
End Function

' Close book.
' Before call this function, should set bellow member variables :
' vpcTarget.sFullPath   : Book of specified path is closed.
' If specified folder path is not exist, this process finish.
' Book is closed return True.
Public Function CloseBook(ByVal vpcTarget As cPathCreator) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    If Not m_BookExistenceChecker.IsBookInOpen(vpcTarget) Then
        CloseBook = False
        Exit Function
    End If
    
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    If Not ThisWorkbook.Name = wbTarget.Name Then
        Application.DisplayAlerts = False
        wbTarget.Close SaveChanges:=True
        Application.DisplayAlerts = True
    End If
    CloseBook = True
End Function

' Copy source file to specified destination file path.
' Before call this function, should set bellow member variables :
' m_pcSource                : Copied file folder path.
' m_pcDestination.sFullPath : Copied file is saved path.
' If specified destination file full path is not exist, process finish.
Public Function CopyBook(ByVal vpcSource As cPathCreator, ByVal vpcDestination As cPathCreator) As Boolean
    If "" <> Dir(vpcDestination.sFullPath) Then
        CopyBook = False
        Exit Function
    End If
    
    Dim bSourceBookOpened As Boolean: bSourceBookOpened = False
    bSourceBookOpened = OpenBook(vpcSource)
    
    Dim wbSource As Workbook
    Set wbSource = Workbooks(vpcSource.sFileName)
    wbSource.SaveCopyAs (vpcDestination.sFullPath)
    Application.DisplayAlerts = False
    If bSourceBookOpened Then: Call CloseBook(vpcSource)
    Application.DisplayAlerts = True
    CopyBook = True
End Function

' Copy source file to specified destination file path.
' Before call this function, should set bellow member variables :
' vpcTarget             : Remove file path.
' vpcTarget.sFullPath   : Remove file full path.
' If specified destination file full path is not exist, process finish.
Public Function RemoveBook(ByVal vpcTarget As cPathCreator) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Or ThisWorkbook.Name = vpcTarget.sFileName Then
        RemoveBook = False
        Exit Function
    End If
    
    If m_BookExistenceChecker.IsBookInOpen(vpcTarget) Then: Call CloseBook(vpcTarget)
    
    Application.DisplayAlerts = False
    Kill (vpcTarget.sFullPath)
    Application.DisplayAlerts = True
    RemoveBook = True
End Function
