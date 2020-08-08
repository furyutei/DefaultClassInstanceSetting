Attribute VB_Name = "Mod_Library"
Option Explicit

Private Enum VBComponentType
    Module = 1
    ClassModule = 2
    Form = 3
    Document = 100
End Enum

Private Enum FsoFolderspec
    WindowsFolder = 0
    SystemFolder = 1
    TemporaryFolder = 2
End Enum

Private Enum ConnectModeEnum
    adModeRead = 1
    adModeWrite = 2
    adModeReadWrite = 3
End Enum

Private Enum StreamTypeEnum
    adTypeBinary = 1
    adTypeText = 2
End Enum

Private Enum StreamReadEnum
    adReadAll = -1
    adReadLine = -2
End Enum

Private Enum StreamWriteEnum
    adWriteChar = 0
    adWriteLine = 1
End Enum

Private Enum SaveOptionsEnum
    adSaveCreateNotExist = 0
    adSaveCreateOverWrite = 2
End Enum


' ■デフォルトクラスインスタンス有無取得
Public Property Get DefaultClsInstance(ClsModuleName As String, Optional TargetBook As Workbook, Optional DebugMode As Boolean = False) As Boolean
    Dim vb_components As Object
    Dim source_component As Object

    If TargetBook Is Nothing Then Set TargetBook = Application.ActiveWorkbook

    Set vb_components = TargetBook.VBProject.VBComponents
    Set source_component = vb_components(ClsModuleName)

    Select Case source_component.Type
        Case VBComponentType.Document
        Case VBComponentType.Form
        Case VBComponentType.ClassModule
        Case Else
            Exit Property
    End Select

    Dim export_filepath As String

    With CreateObject("Scripting.FileSystemObject")
        export_filepath = .BuildPath(.GetSpecialFolder(FsoFolderspec.TemporaryFolder), .GetTempName & ".cls")
    End With

    If DebugMode Then Debug.Print "エクスポート先: " & export_filepath

    source_component.Export export_filepath
    
    Dim code_buffer As String
    Dim ado_stream As Object

    Set ado_stream = CreateObject("ADODB.Stream")
    With ado_stream
        .Mode = ConnectModeEnum.adModeReadWrite
        .Type = StreamTypeEnum.adTypeText

        '参考: [ADODB.StreamのCharsetプロパティに設定できる値 - Jikoryuu’s BLOG](https://jikoryuu.hatenablog.com/entry/66676516)
        '.Charset = "shift-jis"
        .Charset = "x-ms-cp932"

        .Open

        .LoadFromFile export_filepath
        code_buffer = .ReadText(StreamReadEnum.adReadAll)
                
        .Close

        If DebugMode Then Debug.Print "コード:" & vbCrLf & code_buffer
    
        With CreateObject("VBScript.RegExp")
            .Global = False
            .MultiLine = True
            .IgnoreCase = True
    
            .Pattern = "^\s*Attribute\s*VB_PredeclaredId\s*=\s*True$"
            
            Dim matches

            Set matches = .Execute(code_buffer)
            If 0 < matches.Count Then DefaultClsInstance = True
        End With
    End With

    If DebugMode Then
        ' デバッグモード時は一時ファイルを削除しない
    Else
        Kill export_filepath
    End If
End Property


' ■デフォルトクラスインスタンス有無変更
Public Property Let DefaultClsInstance(ClsModuleName As String, Optional TargetBook As Workbook, Optional DebugMode As Boolean = False, IsValid As Boolean)
    Dim vb_components As Object
    Dim source_component As Object

    If TargetBook Is Nothing Then Set TargetBook = Application.ActiveWorkbook
    If TargetBook.Name = ThisWorkbook.Name Then Exit Property

    Set vb_components = TargetBook.VBProject.VBComponents
    Set source_component = vb_components(ClsModuleName)

    If source_component.Type <> VBComponentType.ClassModule Then
        ' 対応するのはクラスモジュールのみ
        Err.Raise 5, Description:="Module type is incorrect"
    End If

    Dim export_filepath As String

    With CreateObject("Scripting.FileSystemObject")
        export_filepath = .BuildPath(.GetSpecialFolder(FsoFolderspec.TemporaryFolder), .GetTempName & ".cls")
    End With

    If DebugMode Then Debug.Print "エクスポート先: " & export_filepath

    source_component.Export export_filepath
    vb_components.Remove source_component
    
    Dim code_buffer As String
    Dim ado_stream As Object

    Set ado_stream = CreateObject("ADODB.Stream")
    With ado_stream
        .Mode = ConnectModeEnum.adModeReadWrite
        .Type = StreamTypeEnum.adTypeText

        '参考: [ADODB.StreamのCharsetプロパティに設定できる値 - Jikoryuu’s BLOG](https://jikoryuu.hatenablog.com/entry/66676516)
        '.Charset = "shift-jis"
        .Charset = "x-ms-cp932"

        .Open

        .LoadFromFile export_filepath
        code_buffer = .ReadText(StreamReadEnum.adReadAll)
        
        .Close

        If DebugMode Then Debug.Print "変更前:" & vbCrLf & code_buffer
    
        With CreateObject("VBScript.RegExp")
            .Global = False
            .MultiLine = True
            .IgnoreCase = True
    
            .Pattern = "^\s*Attribute\s*VB_PredeclaredId\s*=.*$"
            code_buffer = .Replace(code_buffer, "")
            
            .Pattern = "^(\s*Attribute\s*VB_Name\s*=.*)$"
            code_buffer = .Replace(code_buffer, "$1" & "Attribute VB_PredeclaredId = " & IIf(IsValid, "True", "False"))

            .MultiLine = False
            .Pattern = "\s+$"
            code_buffer = .Replace(code_buffer, vbCrLf)
        End With
    
        If DebugMode Then Debug.Print "変更後:" & vbCrLf & code_buffer

        .Open

        '.Position = 0
        .WriteText code_buffer, StreamWriteEnum.adWriteChar
        .SaveToFile export_filepath, SaveOptionsEnum.adSaveCreateOverWrite
                
        .Close
    End With

    vb_components.Import export_filepath
    If DebugMode Then
        ' デバッグモード時は一時ファイルを削除しない
    Else
        Kill export_filepath
    End If
End Property
