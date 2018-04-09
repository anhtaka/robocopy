@echo off

rem 日付を取得
set YYYYMMDD=%DATE:/=%

rem コピー元、コピー先、ログファイル名の設定ファイル呼び出し
FOR /F "tokens=1,2,3 delims=/" %%A IN (robocopy_bkup_conf.txt) DO (
   rem VBスクリプトを実行 %%A:バックアップ元 %%B:バックアップ先 %%C:ログファイル名
   cscript "robocopy.vbs" %%A %%B %%ClogrobocopyLog_%YYYYMMDD%.txt
)

rem 古いログファイル削除
rem cscript "LogOldFile_Del.vbs"