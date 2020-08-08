Attribute VB_Name = "Mod_Form"
Option Explicit

'--------------------------------------------------------------------------------
' �Q�l�F[EXCEL VBA���� - ���[�U�[�t�H�[������ɍőO�ʂɂ���(Excel2016) - hake�̓��L](https://hake.hatenablog.com/entry/20180318/p1)
' SetWindowPos() / FindWindow() �̒�`
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

#If VBA7 Then
    Public Declare PtrSafe Function SetWindowPos Lib "user32" _
        (ByVal hWnd As LongPtr, _
            ByVal hWndInsertAfter As LongPtr, _
            ByVal x As LongPtr, _
            ByVal y As LongPtr, _
            ByVal cx As LongPtr, _
            ByVal cy As LongPtr, _
            ByVal uFlags As LongPtr) As Long
    
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Long
#Else
    Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal x As Long, _
            ByVal y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Long
#End If
'--------------------------------------------------------------------------------

Private Enum VBComponentType
    Module = 1
    ClassModule = 2
    Form = 3
    Document = 100
End Enum

Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

Function Form_GetCurrentWorkbookName() As String
    With FormDefaultClsInstanceSetting.ComboBox_Workbooks
        Form_GetCurrentWorkbookName = .List(.ListIndex)
    End With
End Function

Function Form_GetCurrentWorkbook() As Workbook
    Set Form_GetCurrentWorkbook = Application.Workbooks(Form_GetCurrentWorkbookName)
End Function

Function Form_GetCurrentClassModuleName() As String
    With FormDefaultClsInstanceSetting.ComboBox_ClassModules
        Form_GetCurrentClassModuleName = .List(.ListIndex)
    End With
End Function

Sub Form_Workbook_OnSelected(WorkbookName As String)
    Dim current_book As Workbook

    Set current_book = Application.Workbooks(WorkbookName)

    Dim current_component As Object
    
    FormDefaultClsInstanceSetting.IsUpdating = True

    With FormDefaultClsInstanceSetting.ComboBox_ClassModules
        .Clear

        For Each current_component In current_book.VBProject.VBComponents
            If current_component.Type = VBComponentType.ClassModule Then
                .AddItem current_component.Name
            End If
        Next current_component

        If 0 < .ListCount Then
            .ListIndex = 0
            Form_ClassModule_OnSelected .List(.ListIndex)
        End If

        .SetFocus
    End With

    FormDefaultClsInstanceSetting.IsUpdating = False
End Sub

Sub Form_ClassModule_OnSelected(ClassModuleName As String)
    ' ���݂̃f�t�H���g�C���X�^���X��Ԃ�\��
    With FormDefaultClsInstanceSetting
        If DefaultClsInstance(ClassModuleName, Form_GetCurrentWorkbook, FormDefaultClsInstanceSetting.DebugMode) Then
            .OptionButton_DefaultInstance_Enabled = True
            .OptionButton_DefaultInstance_Disabled = False
        Else
            .OptionButton_DefaultInstance_Enabled = False
            .OptionButton_DefaultInstance_Disabled = True
        End If
    End With
End Sub

Sub ShowDefaultClsInstanceSettingForm(Optional FromIDE As Boolean = False)
    Dim current_book As Workbook
    Dim active_book_name As String

    active_book_name = Application.ActiveWorkbook.Name

    With FormDefaultClsInstanceSetting
        .IsUpdating = True
    
        With .ComboBox_Workbooks
            .Clear
    
            For Each current_book In Application.Workbooks
                .AddItem current_book.Name
                
                If active_book_name = current_book.Name Then
                    .ListIndex = .ListCount - 1
                    Form_Workbook_OnSelected active_book_name
                End If
            Next current_book
        End With

        .IsUpdating = False

        ' TODO: �t�H�[���݂̂�O�ʂɏo���i�u�b�N��O�ʂɏo���Ȃ��j�悤�ɂ��������������킩��Ȃ�
        ' TODO: VBE �̃��j���[�̕�����N������ƁA�N�������� VBE �Ƀt�H�[�J�X���߂��Ă��܂������@���킩��Ȃ�
        ' �� FindWindow() / SetWindowPos �g�p�ɂ�胆�[�U�[�t�H�[�����őO�ʂɒu�����ƂőΉ��\�Ȗ͗l
        '   �Q�l�F[EXCEL VBA���� - ���[�U�[�t�H�[������ɍőO�ʂɂ���(Excel2016) - hake�̓��L](https://hake.hatenablog.com/entry/20180318/p1)

        'Application.Windows(active_book_name).ActivateNext
        'Application.Windows(active_book_name).Activate
        'Application.Windows(active_book_name).WindowState = xlMinimized
        'Application.Visible = True
        'AppActivate Application.Caption

        If FromIDE Then
            'Application.Visible = False ' Modeless�ŋN������K�v�����邽�߁ATrue�ɖ߂��^�C�~���O���킩��Ȃ�
            Application.Windows(active_book_name).WindowState = xlMinimized

            '.Show vbModal
            .Show vbModeless
            
            Dim ret As Long
            Dim formHWnd As Long
        
            'Get window handle of the userform
            formHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, FormDefaultClsInstanceSetting.Caption)
            If formHWnd = 0 Then Debug.Print Err.LastDllError
        
            'Set userform window to 'always on top'
            ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
            If ret = 0 Then Debug.Print Err.LastDllError

            'Application.Visible = True
        Else
            .Show vbModal
        End If
    End With
End Sub

Public Sub Main()
    ShowDefaultClsInstanceSettingForm
End Sub

