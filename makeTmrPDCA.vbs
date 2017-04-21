'WshShell オブジェクト
Set objWshShell = WScript.CreateObject("WScript.Shell")
Set objFS = CreateObject("Scripting.FileSystemObject")

'外部ファイルを読込
Set objWSH_FUNC = objFS.OpenTextFile(".\getDayFile.vbs")
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close

'曜日取得(金曜日なら月曜日のファイルを作ろうとしてるが未実装)
Dim lngWeekday
lngWeekday = Weekday(Now)

'ファイルを上書きコピーする
Dim str_from
str_from = objWshShell.CurrentDirectory & "\" & "hinagata.md"
Dim str_to '一つ上のフォルダとする
str_to   = GetDayFile(1)

If objFS.FileExists(str_to) = False Then
    Call objFS.CopyFile(str_from, str_to)
End If

Set objWshShell = Nothing
Set objWSH_FUNC = Nothing
Set objFS = Nothing

'ファイルオープン
OpenFileBySpecificApp(str_to)
