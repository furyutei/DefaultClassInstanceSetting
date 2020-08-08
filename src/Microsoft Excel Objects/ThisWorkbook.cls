VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' 参考: [Excel アドインを VBA で作成する?|?knooto](https://knooto.info/excel-vba-addin/)
Private Const ADDIN_MENU_NAME = "Default Class Instance Setting"
Private Const IDE_MENU_NAME = "Pre-Instance(&P)" ' TODO: (&P)(VB_PredeclaredIdの"P")→キー被りした場合どうするか？
Private Const CALL_PROC = "Mod_Form.Main"

Private Sub Workbook_Open()
    AddMenus
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DeleteMenus
End Sub

Private Sub AddMenus()
    AddMenuToWorksheetToolbar
    AddMenuToIdeToolbar
End Sub

Private Sub DeleteMenus()
    DeleteMenuFromWorksheetToolbar
    DeleteMenuFromIdeToolbar
End Sub

Private Sub AddMenuToWorksheetToolbar()
    DeleteMenuFromWorksheetToolbar ' 既存のものが残っている場合は削除

    With Application.CommandBars("Worksheet Menu Bar")
        With .Controls.Add(Type:=msoControlButton, Before:=.Controls.Count, Temporary:=True)
            .Caption = ADDIN_MENU_NAME
            .Style = msoButtonIconAndCaption ' [MsoButtonStyle 列挙 (Office) | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle)
            .FaceId = 220 ' https://www.microsoft.com/en-us/download/details.aspx?id=50745
            .OnAction = CALL_PROC
        End With
   End With
End Sub

Private Sub DeleteMenuFromWorksheetToolbar()
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls(ADDIN_MENU_NAME).Delete
      On Error GoTo 0
   End With
End Sub

Private Sub AddMenuToIdeToolbar()
    DeleteMenuFromIdeToolbar ' 既存のものが残っている場合は削除

    Static command_button As ClsIdeCommandButton
    Dim ctrl As CommandBarControl

    Set command_button = New ClsIdeCommandButton
    Set ctrl = command_button.Init(GetIdeMenuBar, IDE_MENU_NAME, CALL_PROC).Control

    ctrl.Visible = True
End Sub

Private Sub DeleteMenuFromIdeToolbar()
    Dim ctrl As CommandBarControl

    For Each ctrl In GetIdeMenuBar.Controls
        If ctrl.Caption = IDE_MENU_NAME Then
            On Error Resume Next
            ctrl.Delete
            On Error GoTo 0
        End If
    Next ctrl
End Sub

Private Function GetIdeMenuBar() As CommandBar
    Dim cb As CommandBar
    
    For Each cb In Application.VBE.CommandBars
        If cb.Type = msoBarTypeMenuBar Then
            Set GetIdeMenuBar = cb
            Exit For
        End If
    Next
End Function
