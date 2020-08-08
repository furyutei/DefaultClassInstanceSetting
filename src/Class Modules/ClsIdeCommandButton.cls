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

' �Q�l�F[VBE���j���[�ɃC�x���g�ɔ�������{�^����ǉ�����](https://gist.github.com/baoo777/5331178)
' �uMicrosoft Visual Basic for Applications Extensibility 5.3�v�ւ̎Q�Ɛݒ肪�K�v

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
    Application.Run ctrl.OnAction
End Sub