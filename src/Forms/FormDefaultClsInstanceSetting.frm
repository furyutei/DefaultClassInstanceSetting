VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormDefaultClsInstanceSetting 
   Caption         =   "Default Class Instance Setting"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "FormDefaultClsInstanceSetting.frx":0000
   StartUpPosition =   2  '画面の中央
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
    If DebugMode Then Debug.Print "ワークブック名が変わりました: " & Form_GetCurrentWorkbookName
    Form_Workbook_OnSelected Form_GetCurrentWorkbookName
End Sub

Private Sub ComboBox_ClassModules_Change()
    If IsUpdating Then Exit Sub
    Form_ClassModule_OnSelected Form_GetCurrentClassModuleName
    If DebugMode Then Debug.Print "クラスモジュール名が変わりました: " & Form_GetCurrentClassModuleName
End Sub

Private Sub CommandButton_Cancel_Click()
    If DebugMode Then Debug.Print "キャンセルボタンもしくはEscキーが押されました"
    Unload Me
    Application.Visible = True
End Sub

Private Sub OptionButton_DefaultInstance_Enabled_Click()
    If IsUpdating Then Exit Sub
    If DebugMode Then Debug.Print "Enabled が選択されました"
    DefaultClsInstance(Form_GetCurrentClassModuleName, Form_GetCurrentWorkbook, DebugMode) = True
End Sub

Private Sub OptionButton_DefaultInstance_Disabled_Click()
    If IsUpdating Then Exit Sub
    If DebugMode Then Debug.Print "Disabled が選択されました"
    DefaultClsInstance(Form_GetCurrentClassModuleName, Form_GetCurrentWorkbook, DebugMode) = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.Visible = True
    If DebugMode Then Debug.Print "フォームが閉じられます: CloseMode=" & CStr(CloseMode)
End Sub
