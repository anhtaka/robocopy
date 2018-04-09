Option Explicit

Dim objNetWork ,  objShell
Dim strSubject  
Dim strLocalDir,strRemoteDir,strLogFile
Dim strRobocopyStatus

'必要なオブジェクトをセット
Set objNetWork = WScript.CreateObject("WScript.Network")  'ホスト名、ユーザ名の取得
Set objShell = CreateObject("WScript.Shell")              'ネットワークドライブのマウントとrobocopyコマンド

'robocopy.batから渡された引数をチェックし、正しく取得できなければエラーメールを送信して終了
If WScript.Arguments.Count <> 3 Then
    strSubject = "[Robocopy_BackupError][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName 
    SendMail strSubject , "Host:" & objNetWork.ComputerName & "のバックアップ処理に失敗しました" & vbCrLf & "Error:引数が不正です「robocopy.bat」を確認してください"
End If
 
'引数を取得
strLocalDir = WScript.Arguments.Item(0)   'コピー元ディレクトリ
strRemoteDir = WScript.Arguments.Item(1)  'コピー先ディレクトリ
strLogFile = WScript.Arguments.Item(2)  'ログファイル

WScript.Echo strLocalDir
WScript.Echo strRemoteDir

'robocopyがインストールされているか確認。robocopyを引数なしで実行しエラーとなるかどうかで、コマンドがあるかどうか判断。
On Error Resume Next
strRobocopyStatus = objShell.Run("robocopy.exe",7,True)

If Err.Number <> 0 Then
    On Error Goto 0
    strSubject = "[Robocopy_BackupError][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName
    SendMail strSubject ,  "エラー:robocopyがインストールされていません。" & intErr
    WScript.Quit
End If
On Error Goto 0

 
'バックアップ(第二引数が7なのでウィンドウを最小化したまま。 /TEEオプション付けてるので進捗を表示します。コマンドプロンプト最小化しないと描画に時間取られるので最小化してます)
'strRobocopyStatus = objShell.Run("cmd /c robocopy.exe " & strLocalDir & " " & strRemoteDir & " " & " /E /COPYALL  /R:2 /W:5 /FFT /TEE /NDL /NP /LOG:" & strLogFile ,7,True)
'ミラーバックアップの場合
strRobocopyStatus = objShell.Run("cmd /c robocopy.exe " & strLocalDir & " " & strRemoteDir & " " & " /LOG:" & strLogFile & " /COPY:DAT",7,True)

Dim strMsg , strStatus ,iRobocopyStatus , strInfo
iRobocopyStatus = CInt(strRobocopyStatus)
strStatus = "------------------------------" & vbCrLf  & "Status: " & vbCrLf

strInfo = "バックアップ元:" & strLocalDir & vbCrLf & "バックアップ先:"& strRemoteDir & vbCrLf

'robocopyからの戻り値を論理積を取って詳細メッセージ生成
If (iRobocopyStatus AND 16) = 16 Then
    strStatus = strStatus & "***FATAL ERROR***"
    strMsg = strMsg & "Code16:重大なエラー。Robocopyは，ファイルをひとつもコピーしませんでした。これは，使用法の間違いか転送元または転送先ディレクトリへのアクセス権が不十分なために発生するエラーです。" & vbCrLf 
End If
If (iRobocopyStatus AND 8) = 8 Then
    strStatus = strStatus & " FAIL "
    strMsg = strMsg & "Code8:いくつかのファイルまたはディレクトリがコピーできませんでした（コピーエラーが発生し，リトライ上限を上回りました）。以降のエラーをチェックしてください。" & vbCrLf 
End If
If (iRobocopyStatus AND 4) = 4 Then
    strStatus = strStatus & " MISM "
    strMsg = strMsg & "Code4:いくつかMismatched(転送元のファイルと同名のディレクトリが転送先にある,あるいはその逆)ファイルまたはディレクトリがみつかりました。ログを調べてください。多分掃除が必要です。" & vbCrLf 
End If
If (iRobocopyStatus AND 2) = 2 Then
    strStatus = strStatus & " XTRA "
    strMsg = strMsg & "Code2:いくつかのExtra(転送元に存在しないのに転送先に存在するファイル)ファイルまたはディレクトリがみつかりました。出力ログを調べてください。掃除が必要かもしれません。" & vbCrLf 
End If
If (iRobocopyStatus AND 1) = 1 Then
    strStatus = strStatus & " COPY "
    strMsg = strMsg & "Code1:一つ以上のファイルが，うまくコピーされました（つまり新しいファイルは転送先に届きました）。" & vbCrLf 
End If
If iRobocopyStatus = 0 Then
    strMsg = strMsg & "Code0:エラーは発生せず，コピーもされませんでした。転送元と転送先ディレクトリツリーは，完全に同期しています。" & vbCrLf 
End If

strStatus = strStatus & vbCrLf  & "---------------------------------"

'robocopyからの戻り値によって処理を振り分ける8以上が致命的なエラー
If iRobocopyStatus >= 8 Then
    strSubject = "[Robocopy_BackupError][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName
    SendMail strSubject , "Host:" & objNetWork.ComputerName & "のバックアップ処理に致命的なエラーが発生しました。ログファイルを確認してください。" & vbCrLf & _
    vbCrLf & strStatus & vbCrLf &  strInfo & vbCrLf & vbCrLf & "詳細情報:" & vbCrLf & strMsg
Else
    strSubject = "[Robocopy_BackupLog][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName
    SendMail strSubject , "Host:" & objNetWork.ComputerName & "のバックアップ処理正常に終了しました。" & vbCrLf & _
    vbCrLf & strStatus & vbCrLf &  strInfo & vbCrLf & vbCrLf & "詳細情報:" & vbCrLf & strMsg
End If

Set objNetWork = Nothing

'こういう分け方もできる(MSのドキュメントより)
'if strRobocopyStatus 16  echo  ***FATAL ERROR***  & goto end
'if strRobocopyStatus 15  echo FAIL MISM XTRA COPY & goto end
'if strRobocopyStatus 14  echo FAIL MISM XTRA      & goto end
'if strRobocopyStatus 13  echo FAIL MISM      COPY & goto end
'if strRobocopyStatus 12  echo FAIL MISM           & goto end
'if strRobocopyStatus 11  echo FAIL      XTRA COPY & goto end
'if strRobocopyStatus 10  echo FAIL      XTRA      & goto end
'if strRobocopyStatus  9  echo FAIL           COPY & goto end
'if strRobocopyStatus  8  echo FAIL                & goto end
'if strRobocopyStatus  7  echo      MISM XTRA COPY & goto end
'if strRobocopyStatus  6  echo      MISM XTRA      & goto end
'if strRobocopyStatus  5  echo      MISM      COPY & goto end
'if strRobocopyStatus  4  echo      MISM           & goto end
'if strRobocopyStatus  3  echo           XTRA COPY & goto end
'if strRobocopyStatus  2  echo           XTRA      & goto end
'if strRobocopyStatus  1  echo                COPY & goto end
'if strRobocopyStatus  0  echo    --no change--    & goto end
 
 
'メール送信処理
Function SendMail2(strSubject , strError)
	print strSubject
	print strError
End Function
Function SendMail(strSubject , strError)
    Dim strMailAddrTo , strSMTPServ , strMailAddrFrom , iSMTPPort
    '# メールアドレスを設定
    strMailAddrTo = "tytakayanagi@nissho-ele.co.jp"
    strSMTPServ = "172.25.34.1"
    strMailAddrFrom = "grandit-skc-support@nissho-ele.co.jp"
    iSMTPPort = 25

    Dim objMail
    Set objMail = CreateObject("CDO.Message")
    'メールを送信
    objMail.From = strMailAddrFrom
    objMail.To = strMailAddrTo
    objMail.Subject = strSubject
    objMail.TextBody = strError
    objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServ
    objMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = iSMTPPort
    objMail.Configuration.Fields.Update
    objMail.Send
    Set objMail = Nothing
End Function