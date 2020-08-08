VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMnuEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �Q�Ɛݒ�F[VBE���j���[�ɃC�x���g�ɔ�������{�^����ǉ�����](https://gist.github.com/baoo777/5331178)
'
' /////////////////////////////////////////////////////////////////////////////////////
' //
' //  ���� clsMnuEvent
' //
' //    VBE���j���[�ɔ������邽�߂̃C�x���g�N���X
' //
' //    ������������CommandBarControl�N���X��Ԃ����ƂŊȒP�Ɏg�p�ł���悤�ɂ��Ă���
' //
' //    Ex)�W�����W���[���ɋL�q�BAddMenuControl�����s
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
' //#���� Init
' //
' //#�T�v VBE���j���[�C�x���g�N���X�̏����������s����
' //
' //#���� cb:�Ԃ�CommandBarControl��}������CommandBar
' //      strCaption:�Ԃ�CommandBarControl�̃L���v�V����
' //      strOnAction:�Ԃ�CommandBarControl�������Ɏ��s�����v���V�[�W����
' //
' //#�ߒl CommandBarControl
' //
' //#��� VBE�C�x���g�̃N���X�����A����������CommandBarControl��Ԃ�
' //
' //#���� 2013/04/07
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