'WshShell �I�u�W�F�N�g
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")

'�O���t�@�C����Ǎ�
Set objWSH_FUNC = objFS.OpenTextFile(".\getDayFile.vbs")
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close
 
'�ЂȌ`�t�@�C�����擾
Dim str_from
str_from = objWshShell.CurrentDirectory & "\" & "hinagata.txt"
'�R�s�[��t�@�C�����擾
Dim str_to
str_to   = GetDayFile(0)

'�t�@�C�����܂��Ȃ���΂ЂȌ`����R�s�[
If objFS.FileExists(todayFile) = False Then
    Call objFS.CopyFile(str_from, str_to)
End If

Set objWshShell = Nothing
Set objWSH_FUNC = Nothing
Set objFS = Nothing

'�t�@�C���I�[�v��
CreateObject("Shell.Application").ShellExecute str_to
