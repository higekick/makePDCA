
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

   GetDayFile = objWshShell.CurrentDirectory & "\..\" & strFormattedDay & ".md"
   Set objWshShell = Nothing
End Function

'引数で受けたファイル名を特定のファイルで開く
Function OpenFileBySpecificApp(file)
 'ファイルオープン
 Set obj = WScript.CreateObject("WScript.Shell")
 'ファイルを開くアプリケーションをフルパスで指定
 Dim app
 app = "C:\Users\jortest\AppData\Local\atom\atom.exe"

 obj.Run app & " - " & file
End Function
