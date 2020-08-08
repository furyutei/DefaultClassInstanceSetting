Attribute VB_Name = "Mod_Form"
Option Explicit

Private Enum VBComponentType
    Module = 1
    ClassModule = 2
    Form = 3
    Document = 100
End Enum

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
    ' 現在のデフォルトインスタンス状態を表示
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

Private Sub ShowDefaultClsInstanceSettingForm()
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

        ' TODO: フォームのみを前面に出す（ブックを前面に出さない）ようにしたいがやり方がわからない
        'Application.Windows(active_book_name).ActivateNext
        Application.Windows(active_book_name).Activate
        'Application.Windows(active_book_name).WindowState = xlMinimized
        'Application.Visible = False

        ' TODO: VBE のメニューの方から起動すると、起動した後 VBE にフォーカスが戻ってしまう回避方法がわからない
        '.Show vbModeless
        .Show vbModal

        Application.Visible = True
    End With
End Sub

Public Sub Main()
    ShowDefaultClsInstanceSettingForm
End Sub

