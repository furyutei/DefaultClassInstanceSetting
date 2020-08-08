' �Q�l�F [VBScript �� Excel �ɃA�h�C���������ŃC���X�g�[��/�A���C���X�g�[��������@: ���� SE �̂Ԃ₫](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

On Error Resume Next

Dim installPath
Dim IsJA
Dim addInName
Dim addInFileName
Dim objExcel
Dim objAddin
Dim objWshShell
Dim objFileSys

Function IIf(ByVal str, ByVal trueval, ByVal falseval)
    Dim rtn
    If str Then
        rtn = trueval
    Else
        rtn = falseval
    End If
    IIf = rtn
End Function

IsJA = GetLocale() = 1041

'�A�h�C������ݒ�
addInName = IIf(IsJA, "�f�t�H���g�C���X�^���X�ݒ�", "Default Class Instance Setting")
addInFileName = "DefaultClassInstanceSetting.xlam"

'Excel���쒆����
Err.Clear
Set objExcel = GetObject(, "Excel.Application")
If Err.Number = 0 Then
    Set objExcel = Nothing
    MsgBox IIf(IsJA, "Excel ��S�ĕ��Ă��������I", "Please close all Excel applications !"), vbExclamation,addInName
    WScript.Quit
End If
Err.Clear

IF MsgBox(IIf(IsJA, "�A�h�C�����A���C���X�g�[�����܂����H", "Do you want to uinstall this add-in ?"), vbYesNo + vbQuestion, addInName) = vbNo Then
    WScript.Quit
End IF

'Excel �C���X�^���X��
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'�A�h�C���o�^����
For i = 1 To objExcel.Addins.Count
    Set objAddin = objExcel.Addins.item(i)
    If objAddin.Name = addInFileName Then
        objAddin.Installed = False
    End If
Next

'Excel �I��
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���폜
If objFileSys.FileExists(installPath) Then
    objFileSys.DeleteFile installPath, True
Else
    'MsgBox "�A�h�C���t�@�C�������݂��܂���B", vbExclamation, addInName
End If

Set objWshShell = Nothing
Set objFileSys = Nothing

IF Err.Number = 0 THEN
    MsgBox IIF(IsJA, "�A�h�C���̃A���C���X�g�[�����������܂���", "Uninstallation is now complete !"), vbInformation, addInName
ELSE
    MsgBox IIf(IsJA, "�G���[���������܂���: " & CStr(Err.Number) & vbCrLF & "���s�����m�F���Ă�������", "An error has occurred." & vbCrLF & "Please check your environment."), vbExclamation, addInName
End IF
