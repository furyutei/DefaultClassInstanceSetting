VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDefaultClsInstanceSetting 
   Caption         =   "Default Class Instance Setting"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "FormDefaultClsInstanceSetting.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "FormDefaultClsInstanceSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DebugMode_ As Boolean
Private IsUpdating_ As Boolean

Public Property Get DebugMode() As Boolean
    DebugMode = DebugMode_
End Property

Public Property Let DebugMode(SetValue As Boolean)
    DebugMode_ = SetValue
End Property

Public Property Get IsUpdating() As Boolean
    IsUpdating = IsUpdating_
End Property

Public Property Let IsUpdating(SetValue As Boolean)
    IsUpdating_ = SetValue
End Property

Private Sub ComboBox_Workbooks_Change()
    If IsUpdating Then Exit Sub
    Form_Workbook_OnSelected Form_GetCurrentWorkbookName
End Sub

Private Sub ComboBox_ClassModules_Change()
    If IsUpdating Then Exit Sub
    Form_ClassModule_OnSelected Form_GetCurrentClassModuleName
End Sub

Private Sub CommandButton_Cancel_Click()
    Unload Me
    Application.Visible = True
End Sub

Private Sub OptionButton_DefaultInstance_Disabled_Click()
    If IsUpdating Then Exit Sub
    DefaultClsInstance(Form_GetCurrentClassModuleName, Form_GetCurrentWorkbook, DebugMode) = False
End Sub

Private Sub OptionButton_DefaultInstance_Enabled_Click()
    If IsUpdating Then Exit Sub
    DefaultClsInstance(Form_GetCurrentClassModuleName, Form_GetCurrentWorkbook, DebugMode) = True
End Sub
