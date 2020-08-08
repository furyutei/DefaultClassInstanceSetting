VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsIdeCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 参考：[VBEメニューにイベントに反応するボタンを追加する](https://gist.github.com/baoo777/5331178)
' 「Microsoft Visual Basic for Applications Extensibility 5.3」への参照設定が必要

Private WithEvents IdeCommandButtonEvent As VBIDE.CommandBarEvents
Attribute IdeCommandButtonEvent.VB_VarHelpID = -1

Private ButtonControl_ As CommandBarControl

Public Property Get Control() As CommandBarControl
    Set Control = ButtonControl_
End Property

Public Function Init(cb As CommandBar, Optional strCaption As String, Optional strOnAction As String) As ClsIdeCommandButton
    Dim ctrl As CommandBarControl

    Set ctrl = cb.Controls.Add(msoControlButton)
    
    ctrl.Style = msoButtonCaption
    If strCaption <> "" Then ctrl.Caption = strCaption
    If strOnAction <> "" Then ctrl.OnAction = strOnAction

    Set IdeCommandButtonEvent = Application.VBE.Events.CommandBarEvents(ctrl)
    
    Set ButtonControl_ = ctrl
    Set Init = Me
End Function

Private Sub IdeCommandButtonEvent_Click(ByVal ctrl As Object, handled As Boolean, CancelDefault As Boolean)
    Application.Run ctrl.OnAction, True
End Sub

