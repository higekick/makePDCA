'WshShell �I�u�W�F�N�g
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")

'�O���t�@�C����Ǎ�
Set objWSH_FUNC = objFS.OpenTextFile(".\getDayFile.vbs")
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close

'�j���擾(���j���Ȃ猎�j���̃t�@�C������낤�Ƃ��Ă邪������)
Dim lngWeekday
lngWeekday = Weekday(Now)

'�t�@�C�����㏑���R�s�[����
Dim str_from
str_from = objWshShell.CurrentDirectory & "\" & "hinagata.md"
Dim str_to '���̃t�H���_�Ƃ���
str_to   = GetDayFile(1)

If objFS.FileExists(str_to) = False Then
    Call objFS.CopyFile(str_from, str_to)
End If

Set objWshShell = Nothing
Set objWSH_FUNC = Nothing
Set objFS = Nothing

'�t�@�C���I�[�v��
OpenFileBySpecificApp(str_to)
