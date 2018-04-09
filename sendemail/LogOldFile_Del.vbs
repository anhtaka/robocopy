'logファイルの格納されているフォルダーのパス
Const cFolderPath = "D:\sscwork\00.データ移行\20180414_本番移行\robocopy\log"
'以下の文字列が先頭にあるファイルを削除する
Dim strLogShareFile
strLogShareFile = "robocopyLog_"
'以下の日数経っているものを削除対象とする
Dim iDate
iDate=120

Dim fso, fl
Dim tod, logDate, buf, diff
Set fso = CreateObject("Scripting.FileSystemObject")

tod = CDate(Split(Now, " ")(0))
For Each fl In fso.GetFolder(cFolderPath).Files
    ' strLogShareFileが先頭につくファイルのみ削除対象とする
    If InStr(1 , fl.Name , strLogShareFile , 1 ) = 1 Then
        buf = fso.GetFile(fl.Path).DateLastModified
        logDate = CDate(Split(buf, " ")(0))
        diff = DateDiff("d", logDate, tod)
        If diff > iDate Then
            fl.Delete
        End If
    End If
Next
Set fso = Nothing