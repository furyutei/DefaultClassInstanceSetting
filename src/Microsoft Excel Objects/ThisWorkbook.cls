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

Private Sub Workbook_Open()
    With Application.CommandBars("Worksheet Menu Bar")
        On Error Resume Next
        ' 既存のものが残っている場合は削除
        .Controls(ADDIN_MENU_NAME).Delete
        On Error GoTo 0

        With .Controls.Add(Type:=msoControlButton, before:=.Controls.Count, Temporary:=True)
            .Caption = ADDIN_MENU_NAME
            .Style = msoButtonIconAndCaption ' [MsoButtonStyle 列挙 (Office) | Microsoft Docs](https://docs.microsoft.com/ja-jp/office/vba/api/office.msobuttonstyle)
            .FaceId = 220 ' https://www.microsoft.com/en-us/download/details.aspx?id=50745
            .OnAction = "Mod_Form.Main"
        End With
   End With
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls(ADDIN_MENU_NAME).Delete
      On Error GoTo 0
   End With
End Sub

