VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMnuEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 参照設定：[VBEメニューにイベントに反応するボタンを追加する](https://gist.github.com/baoo777/5331178)
'
' /////////////////////////////////////////////////////////////////////////////////////
' //
' //  名称 clsMnuEvent
' //
' //    VBEメニューに反応するためのイベントクラス
' //
' //    初期化処理でCommandBarControlクラスを返すことで簡単に使用できるようにしている
' //
' //    Ex)標準モジュールに記述。AddMenuControlを実行
' //    '------------------------------------------------------------------------------
' //    Private mnu1 As clsMnuEvent
' //    Private Sub AddMenuControl()
' //
' //        Dim cb As CommandBar
' //        Dim ctl1 As CommandBarControl
' //
' //        Set cb = Application.VBE.CommandBars.Add("VBEMenu", msoBarTop, , True)
' //        Set mnu1 = New clsMnuEvent
' //        Set ctl1 = mnu1.Init(cb, "aaa", "abc")
' //
' //        cb.Visible = True
' //
' //    End Sub
' //    Private Sub DelMenuControl()
' //
' //        Dim ctl As CommandBarControl
' //        Dim cb As CommandBar
' //        For Each cb In Application.VBE.CommandBars
' //            If cb.Name = "VBEMenu" Then
' //                cb.Delete
' //            End If
' //        Next
' //    End Sub
' //    Private Sub abc()
' //        Debug.Print "Sample Test"
' //    End Sub
' //    '------------------------------------------------------------------------------
' //
' //    2013/04/07
' //    Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Option Explicit
Private WithEvents m_MenuEvent As VBIDE.CommandBarEvents
Attribute m_MenuEvent.VB_VarHelpID = -1
' /////////////////////////////////////////////////////////////////////////////////////
' //#名称 Init
' //
' //#概要 VBEメニューイベントクラスの初期化を実行する
' //
' //#引数 cb:返すCommandBarControlを挿入するCommandBar
' //      strCaption:返すCommandBarControlのキャプション
' //      strOnAction:返すCommandBarControl押下時に実行されるプロシージャ名
' //
' //#戻値 CommandBarControl
' //
' //#解説 VBEイベントのクラスだが、初期化時にCommandBarControlを返す
' //
' //#履歴 2013/04/07
' //      Coded by YASUTADA OOBA
' //
' /////////////////////////////////////////////////////////////////////////////////////
Public Function Init(cb As CommandBar, Optional strCaption As String, Optional strOnAction As String) As CommandBarControl

    Dim ctl As CommandBarControl
    Set ctl = cb.Controls.Add(msoControlButton)
    If strCaption <> "" Then
        ctl.Caption = strCaption
    End If
    ctl.Style = msoButtonCaption
    If strOnAction <> "" Then
        ctl.OnAction = strOnAction
    End If
    Set m_MenuEvent = Application.VBE.Events.CommandBarEvents(ctl)
    Set Init = ctl

End Function
Private Sub m_MenuEvent_Click(ByVal ctl As Object, handled As Boolean, CancelDefault As Boolean)
    Application.Run ctl.OnAction
End Sub
