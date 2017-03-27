Set objFS = CreateObject("Scripting.FileSystemObject")

'外部ファイルを読込
Set objWSH_FUNC = objFS.OpenTextFile(".\getDayFile.vbs")
ExecuteGlobal objWSH_FUNC.ReadAll()
objWSH_FUNC.Close

'曜日取得
Dim lngWeekday
lngWeekday = Weekday(Now)
'昨日のファイル名取得
Dim yesterDayFile
yesterDayFile = GetDayFile(-1)

If objFS.FileExists(yesterDayFile) = True Then
    '昨日のファイルをオープン
    CreateObject("Shell.Application").ShellExecute yesterdayFile
elseif lngWeekday = vbMonday Then
    '昨日のファイルが存在しなくて、月曜日なら金曜日のファイルを開く
    Dim friDayFile
    friDayFile = GetDayFile(-3)
    CreateObject("Shell.Application").ShellExecute friDayFile
End If

Set objFS = Nothing
Set objWSH_FUNC = Nothing
