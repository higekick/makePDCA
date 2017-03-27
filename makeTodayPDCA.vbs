'WshShell オブジェクト
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")

'外部ファイルを読込
Set objWSH_FUNC = objFS.OpenTextFile(".\getDayFile.vbs")
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close
 
'ひな形ファイル名取得
Dim str_from
str_from = objWshShell.CurrentDirectory & "\" & "hinagata.txt"
'コピー先ファイル名取得
Dim str_to
str_to   = GetDayFile(0)

'ファイルがまだなければひな形からコピー
If objFS.FileExists(todayFile) = False Then
    Call objFS.CopyFile(str_from, str_to)
End If

Set objWshShell = Nothing
Set objWSH_FUNC = Nothing
Set objFS = Nothing

'ファイルオープン
CreateObject("Shell.Application").ShellExecute str_to
