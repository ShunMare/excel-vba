VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRangeInfo_DR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_rDailyReport As Range
Dim m_rPotentialCustomerKPI As Range
Dim m_rNewKPI As Range
Dim m_rTimeKPI As Range
Dim m_rGoal As Range

'--------------------------------------------------------------------------------
'@brief : getter setter rDailyReport
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get rDailyReport() As Range
    Set rDailyReport = m_rDailyReport
End Property

Public Property Set rDailyReport(ByVal vrDailyReport As Range)
    Set m_rDailyReport = vrDailyReport
End Property

'--------------------------------------------------------------------------------
'@brief : getter setter rPotentialCustomerKPI
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get rPotentialCustomerKPI() As Range
    Set rPotentialCustomerKPI = m_rPotentialCustomerKPI
End Property

Public Property Set rPotentialCustomerKPI(ByVal vrPotentialCustomerKPI As Range)
    Set m_rPotentialCustomerKPI = vrPotentialCustomerKPI
End Property

'--------------------------------------------------------------------------------
'@brief : getter setter rNewKPI
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get rNewKPI() As Range
    Set rNewKPI = m_rNewKPI
End Property

Public Property Set rNewKPI(ByVal vrNewKPI As Range)
    Set m_rNewKPI = vrNewKPI
End Property

'--------------------------------------------------------------------------------
'@brief : getter setter rTimeKPI
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get rTimeKPI() As Range
    Set rTimeKPI = m_rTimeKPI
End Property

Public Property Set rTimeKPI(ByVal vrTimeKPI As Range)
    Set m_rTimeKPI = vrTimeKPI
End Property

'--------------------------------------------------------------------------------
'@brief : getter setter rGoal
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Property Get rGoal() As Range
    Set rGoal = m_rGoal
End Property

Public Property Set rGoal(ByVal vrGoal As Range)
    Set m_rGoal = vrGoal
End Property

'--------------------------------------------------------------------------------
'@brief : set number range
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Sub SetNumberRange(ByVal vlTargetNumber As Long, ByVal vrTarget As Range)
    Select Case vlTargetNumber
    Case E_LARGE_ITEM.DAIRY_REPORT
        Set rDailyReport = vrTarget
    Case E_LARGE_ITEM.POTENTIAL_COSTOMER_KPI
        Set rPotentialCustomerKPI = vrTarget
    Case E_LARGE_ITEM.NEW_KPI
        Set rNewKPI = vrTarget
    Case E_LARGE_ITEM.TIME_KPI
        Set rTimeKPI = vrTarget
    Case E_LARGE_ITEM.GOAL
        Set rGoal = vrTarget
    End Select
End Sub

'--------------------------------------------------------------------------------
'@brief : return number range
'@param : NULL
'@return: NULL
'--------------------------------------------------------------------------------
Public Function ReturnNumberRange(ByVal vlTargetNumber As Long) As Range
    Select Case vlTargetNumber
    Case E_LARGE_ITEM.DAIRY_REPORT
        Set ReturnNumberRange = rDailyReport
    Case E_LARGE_ITEM.POTENTIAL_COSTOMER_KPI
        Set ReturnNumberRange = rPotentialCustomerKPI
    Case E_LARGE_ITEM.NEW_KPI
        Set ReturnNumberRange = rNewKPI
    Case E_LARGE_ITEM.TIME_KPI
        Set ReturnNumberRange = rTimeKPI
    Case E_LARGE_ITEM.GOAL
        Set ReturnNumberRange = rGoal
    End Select
End Function


