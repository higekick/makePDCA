
'今日から引数で受けた日付をaddした形式のファイル名を返す
Function GetDayFile(addDay)
   'WshShell オブジェクト
   Dim objWshShell
   Set objWshShell = WScript.CreateObject("WScript.Shell")

   'yyyymmdd 形式で現在日付を取得
   Dim strFormattedToday
   strFormattedToday = Left(Now(),10)

   Dim strFormattedDay
   strFormattedDay = DateAdd("d", addDay, strFormattedToday)
   strFormattedDay = Replace(strFormattedDay, "/", "")

   GetDayFile = objWshShell.CurrentDirectory & "\..\" & strFormattedDay & ".txt"
   Set objWshShell = Nothing
End Function
