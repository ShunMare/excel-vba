Attribute VB_Name = "bGlobal"
Option Explicit

Public Enum E_SEPARATE_STRING
    NORMAL
    BETWEEN
    SEPARATE_STRING_NUM
End Enum

Public Enum E_LARGE_ITEM
    DAIRY_REPORT
    POTENTIAL_COSTOMER_KPI
    NEW_KPI
    TIME_KPI
    GOAL
    LARGE_ITEM_NUM
End Enum

Public Enum E_DIVISION
    DIV_1
    DIV_2
    DIV_3
    DIV_4
    DIV_ALL
    DIV_NUM
End Enum

Public Enum E_EXTENSION
    ALL
    XLS
    XLSX
    XLSM
    TXT
    PDF
End Enum

Public Const G_S_NOT_EXIST As String = "NOT_EXIST"
Public Const G_S_SHEET_NAME_CONFIG As String = "config"

''config
Public Const G_S_CONFIG_DR_ORI_EXPORT_PATH As String = "CONFIG_DR_ORI_EXPORT_PATH"
Public Const G_S_CONFIG_LOCAL_DRIVE As String = "CONFIG_LOCAL_DRIVE"

Public g_siConfig As cSheetInfo
Public g_Config As cConfig
Public g_bTargetDivision(E_DIVISION.DIV_NUM) As Boolean

