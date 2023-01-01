VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSheetOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class is to operate Excel file.
' Can create, show, hide, copy and remove.
'
' When use this class, should import bellow class:
' - cPathCreator
' - cBookExistenceChecker
' - cBookOperator

Private m_BookExistenceChecker As cBookExistenceChecker
Private m_SheetExistenceChecker As cSheetExistenceChecker
Private m_BookOperator As cBookOperator

' Initialize member variables.
Private Sub Class_Initialize()
    Set m_BookExistenceChecker = New cBookExistenceChecker
    Set m_SheetExistenceChecker = New cSheetExistenceChecker
    Set m_BookOperator = New cBookOperator
End Sub

' Check member variables.
' Judge whether to proceed to the next process.
Private Function checkMyValue()
    checkMyValue = True
End Function

' Copy sheet in source book to destination book .
' Before call this function, should set bellow member variables :
' m_sCopySheetName          : Copied sheet name in source book.
' If specified destination book full path is not exist, process finish.
' If m_sChangeSheetName is set, change copied sheet name the value.
Public Function CreateSheet(ByVal vpcTarget As cPathCreator, ByVal vsSheetName As String) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    Call m_BookOperator.OpenBook(vpcTarget)

    Dim wbTarget As Workbook
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    If m_SheetExistenceChecker.HasSheet(vpcTarget, vsSheetName) Then
        Call m_BookOperator.CloseBook(vpcTarget)
        CreateSheet = False
        Exit Function
    End If
    wbTarget.Worksheets.Add before:=Worksheets(1)
    wbTarget.Worksheets(1).Name = vsSheetName
    
    Call m_BookOperator.CloseBook(vpcTarget)
    CreateSheet = True
End Function

' Show sheet in specified book path.
' Before call this function, should set bellow member variables :
' m_sCopySheetName          : Copied sheet name in source book.
' If specified destination book full path is not exist, process finish.
Public Function ShowSheet(ByVal vpcTarget As cPathCreator, ByVal vsSheetName As String) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    Call m_BookOperator.OpenBook(vpcTarget)

    Dim wbTarget As Workbook
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    If Not m_SheetExistenceChecker.HasSheet(vpcTarget, vsSheetName) _
    Or m_SheetExistenceChecker.IsSheetVisible(vpcTarget, vsSheetName) Then
        Call m_BookOperator.CloseBook(vpcTarget)
        ShowSheet = False
        Exit Function
    End If
    wbTarget.Worksheets(vsSheetName).Visible = True
    
    Call m_BookOperator.CloseBook(vpcTarget)
    ShowSheet = True
End Function

' Hide target book in target book path.
' Before call this function, should set bellow member variables :
' m_sTargetSheetName        : Show sheet name in target book.
' If specified target book full path is not exist, process finish.
Public Function HideSheet(ByVal vpcTarget As cPathCreator, ByVal vsSheetName As String)
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    
    Call m_BookOperator.OpenBook(vpcTarget)

    Dim wbTarget As Workbook
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    If Not m_SheetExistenceChecker.HasSheet(vpcTarget, vsSheetName) _
    Or Not m_SheetExistenceChecker.IsSheetVisible(vpcTarget, vsSheetName) Then
        HideSheet = False
        Exit Function
    End If
    wbTarget.Worksheets(vsSheetName).Visible = False
    
    Call m_BookOperator.CloseBook(vpcTarget)
    HideSheet = True
End Function

' Copy source file to specified destination file path.
' Before call this function, should set bellow member variables :
' m_sCopySheetName          : Copied sheet name in source file.
' If specified destination file full path is not exist, process finish.
' If m_sChangeSheetName is set, change copied sheet name the value.
Public Function CopySheet(ByVal vpcSource As cPathCreator, ByVal vpcDestination As cPathCreator _
, ByVal vsSheetName As String) As Boolean
    Call m_BookOperator.OpenBook(vpcSource)
    Call m_BookOperator.OpenBook(vpcDestination)
    
    If Not m_SheetExistenceChecker.HasSheet(vpcSource, vsSheetName) _
    Or m_SheetExistenceChecker.HasSheet(vpcDestination, vsSheetName) Then
        CopySheet = False
        Exit Function
    End If
    
    Dim wbSource As Workbook
    Dim wbDestination As Workbook
    Set wbSource = Workbooks(vpcSource.sFileName)
    Set wbDestination = Workbooks(vpcDestination.sFileName)
    wbSource.Worksheets(vsSheetName).Copy before:=wbDestination.Worksheets(1)
    
    Call m_BookOperator.CloseBook(vpcSource)
    Call m_BookOperator.CloseBook(vpcDestination)
    CopySheet = True
End Function

' Copy source file to specified destination file path.
' Before call this function, should set bellow member variables :
' m_sCopySheetName          : Copied sheet name in source file.
' If specified destination file full path is not exist, process finish.
' If m_sChangeSheetName is set, change copied sheet name the value.
Public Function ChangeSheetName(ByVal vpcTarget As cPathCreator, _
ByVal vsPreSheetName As String, ByVal vsPostSheetName As String) As Boolean
    Call m_BookOperator.OpenBook(vpcTarget)
    
    If Not m_SheetExistenceChecker.HasSheet(vpcTarget, vsPreSheetName) _
    Or m_SheetExistenceChecker.HasSheet(vpcTarget, vsPostSheetName) Then
        ChangeSheetName = False
        Exit Function
    End If
    
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    wbTarget.Worksheets(vsPreSheetName).Name = vsPostSheetName
    
    Call m_BookOperator.CloseBook(vpcTarget)
    ChangeSheetName = True
End Function

' Remove source book in target file.
' Before call this function, should set bellow member variables :
' m_sTargetSheetName        : Remove sheet in target file.
' If specified target is not exist, process finish.
Public Function RemoveSheet(ByVal vpcTarget As cPathCreator, ByVal vsSheetName As String) As Boolean
    If "" = Dir(vpcTarget.sFullPath) Then: Exit Function
    Call m_BookOperator.OpenBook(vpcTarget)

    If Not m_SheetExistenceChecker.HasSheet(vpcTarget, vsSheetName) Then
        RemoveSheet = False
        Exit Function
    End If
    
    Dim wbTarget As Workbook
    Set wbTarget = Workbooks(vpcTarget.sFileName)
    Application.DisplayAlerts = False
    wbTarget.Worksheets(vsSheetName).Delete
    Application.DisplayAlerts = True
    
    Call m_BookOperator.CloseBook(vpcTarget)
    RemoveSheet = True
End Function

