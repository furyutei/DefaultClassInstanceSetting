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

IF MsgBox(IIf(IsJA, "�A�h�C�����C���X�g�[�����܂����H", "Do you want to install this add-in ?"), vbYesNo + vbQuestion, addInName) = vbNo Then
    WScript.Quit
End IF

'Excel �C���X�^���X��
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'�o�^�ςݑ΍�
For i = 1 To objExcel.Addins.Count
    Set objAddin = objExcel.Addins.item(i)
    If objAddin.Name = addInFileName Then
        objAddin.Installed = False
    End If
Next

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'�C���X�g�[����p�X�̍쐬
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'�t�@�C���R�s�[(�㏑��)
objFileSys.CopyFile addInFileName , installPath, True

Set objWshShell = Nothing
Set objFileSys = Nothing

'�A�h�C�� Workbook �^�C�g���ݒ�i����ɉ����ă^�C�g���ύX�j
Set objWorkbook = objExcel.Workbooks.Open(installPath)
objExcel.DisplayAlerts = False
objWorkbook.Title = addInName
objWorkbook.Save
objWorkbook.Close
objExcel.DisplayAlerts = True

'�A�h�C���o�^
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

'Excel �I��
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

IF Err.Number = 0 THEN
    MsgBox IIf(IsJA, "�A�h�C���̃C���X�g�[�����������܂���", "Installation is now complete !"), vbInformation, addInName
ELSE
    MsgBox IIf(IsJA, "�G���[���������܂���: " & CStr(Err.Number) & vbCrLF & "���s�����m�F���Ă�������", "An error has occurred." & vbCrLF & "Please check your environment."), vbExclamation, addInName
End IF
