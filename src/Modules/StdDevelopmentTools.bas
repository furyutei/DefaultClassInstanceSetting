Attribute VB_Name = "StdDevelopmentTools"
Option Explicit

Public Sub ExportAll()
    ConvertToAddIn
    ExportVisualBasicCode
    'ExportCustumUI_Xml
End Sub

Public Sub ConvertToAddIn()
    Dim SourceDirectory As String
    Dim SourceFilename As String
    Dim SourceFilepath As String
    Dim TargetDirectory As String
    Dim TargetFilename As String
    Dim TargetFilepath As String
    Dim Fso As Object: Set Fso = CreateObject("Scripting.FileSystemObject")

    SourceDirectory = ThisWorkbook.Path
    SourceFilename = ThisWorkbook.Name
    SourceFilepath = ThisWorkbook.FullName

    TargetDirectory = Fso.GetAbsolutePathName(Fso.BuildPath(SourceDirectory, "..\addin"))

    If Not Fso.FolderExists(TargetDirectory) Then
        Call Fso.CreateFolder(TargetDirectory)
    End If
    
    TargetFilename = Replace(SourceFilename, ".xlsm", ".xlam")
    TargetFilepath = Fso.BuildPath(TargetDirectory, TargetFilename)
    
    If Dir(TargetFilepath) <> "" Then Kill TargetFilepath
    
    ThisWorkbook.RemovePersonalInformation = True
    ThisWorkbook.RemoveDocumentInformation xlRDIDocumentProperties
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=TargetFilepath, FileFormat:=xlOpenXMLAddIn

    '[�o��]
    '  �A�h�C���Ƃ��ĕۑ���A���t�@�C����ۑ����悤�Ƃ���ƁA
    '  �u�t�@�C�� '...' �́A�O��ۑ����ꂽ��A�ق��̃��[�U�[�ɂ���ĕύX���ꂽ�\��������܂��B�����I�����ĉ������B�v
    '  �̂悤�Ȋm�F�_�C�A���O���o�邽�߁A�����}�����邽�߂ɏ㏑���ۑ����Ă���
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True

    Debug.Print "Converted to " & TargetFilepath
End Sub


'[Excel macro to export all VBA source code in this project to text files for proper source control versioning](https://gist.github.com/steve-jansen/7589478)
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim Count As Integer
    Dim Path As String
    Dim Relative_directory As String
    Dim Directory As String
    Dim Extension As String
    Dim Fso As Object: Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Count = 0
    
    For Each VBComponent In ThisWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case Document
                Extension = ".cls"
                Relative_directory = "Microsoft Excel Objects"
            Case Form
                Extension = ".frm"
                Relative_directory = "Forms"
            Case Module
                Extension = ".bas"
                Relative_directory = "Modules"
            Case ClassModule
                Extension = ".cls"
                Relative_directory = "Class Modules"
            Case Else
                Debug.Print "Type: " & CStr(VBComponent.Type)
                Extension = ".txt"
                Relative_directory = "Others"
        End Select
        
        On Error Resume Next
        Err.Clear
        
        Directory = Fso.BuildPath(ThisWorkbook.Path, Relative_directory)

        If Not Fso.FolderExists(Directory) Then
            Call Fso.CreateFolder(Directory)
        End If
                
        Path = Fso.BuildPath(Directory, VBComponent.Name & Extension)
        Call VBComponent.Export(Path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & Path, vbCritical)
        Else
            Count = Count + 1
            Debug.Print "Exported " & Left(VBComponent.Name & Space(Padding), Padding - 3) & "-> " & Path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(Count) & " VBA files to " & Directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub

Sub ExportCustumUI_Xml()
    Const CustumUI_Directory = "customUI"
    Const CustumUI_Firename = "customUI14.xml"
    Const TempZipFilename = "temp.zip"

    Dim Fso As Object: Set Fso = CreateObject("Scripting.FileSystemObject")
    Dim App As Object: Set App = CreateObject("Shell.Application")
    Dim BaseDirectory As String
    Dim TargetDirectory As String
    Dim TempZipFilepath As String

    BaseDirectory = ThisWorkbook.Path
    TargetDirectory = Fso.BuildPath(BaseDirectory, CustumUI_Directory)

    If Not Fso.FolderExists(TargetDirectory) Then
        Fso.CreateFolder TargetDirectory
    End If
    
    TempZipFilepath = Fso.BuildPath(BaseDirectory, TempZipFilename)
    Fso.CopyFile ThisWorkbook.FullName, TempZipFilepath, True
   
    Dim NamespaceSource As Object
    Dim NamespaceTarget As Object

    Set NamespaceSource = App.Namespace(TempZipFilepath & "\" & CustumUI_Directory)
    Set NamespaceTarget = App.Namespace(TargetDirectory & "\")

    NamespaceTarget.CopyHere NamespaceSource.items.Item(CustumUI_Firename), &H10 ' FOF_NOCONFIRMATION(&H10)

    Kill TempZipFilepath

    Debug.Print "Exported to " & Fso.BuildPath(TargetDirectory, CustumUI_Firename)
End Sub

Sub ClearStatusBar()
    Application.StatusBar = False
End Sub
