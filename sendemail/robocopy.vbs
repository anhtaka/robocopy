Option Explicit

Dim objNetWork ,  objShell
Dim strSubject  
Dim strLocalDir,strRemoteDir,strLogFile
Dim strRobocopyStatus

'�K�v�ȃI�u�W�F�N�g���Z�b�g
Set objNetWork = WScript.CreateObject("WScript.Network")  '�z�X�g���A���[�U���̎擾
Set objShell = CreateObject("WScript.Shell")              '�l�b�g���[�N�h���C�u�̃}�E���g��robocopy�R�}���h

'robocopy.bat����n���ꂽ�������`�F�b�N���A�������擾�ł��Ȃ���΃G���[���[���𑗐M���ďI��
If WScript.Arguments.Count <> 3 Then
    strSubject = "[Robocopy_BackupError][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName 
    SendMail strSubject , "Host:" & objNetWork.ComputerName & "�̃o�b�N�A�b�v�����Ɏ��s���܂���" & vbCrLf & "Error:�������s���ł��urobocopy.bat�v���m�F���Ă�������"
End If
 
'�������擾
strLocalDir = WScript.Arguments.Item(0)   '�R�s�[���f�B���N�g��
strRemoteDir = WScript.Arguments.Item(1)  '�R�s�[��f�B���N�g��
strLogFile = WScript.Arguments.Item(2)  '���O�t�@�C��

WScript.Echo strLocalDir
WScript.Echo strRemoteDir

'robocopy���C���X�g�[������Ă��邩�m�F�Brobocopy�������Ȃ��Ŏ��s���G���[�ƂȂ邩�ǂ����ŁA�R�}���h�����邩�ǂ������f�B
On Error Resume Next
strRobocopyStatus = objShell.Run("robocopy.exe",7,True)

If Err.Number <> 0 Then
    On Error Goto 0
    strSubject = "[Robocopy_BackupError][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName
    SendMail strSubject ,  "�G���[:robocopy���C���X�g�[������Ă��܂���B" & intErr
    WScript.Quit
End If
On Error Goto 0

 
'�o�b�N�A�b�v(��������7�Ȃ̂ŃE�B���h�E���ŏ��������܂܁B /TEE�I�v�V�����t���Ă�̂Ői����\�����܂��B�R�}���h�v�����v�g�ŏ������Ȃ��ƕ`��Ɏ��Ԏ����̂ōŏ������Ă܂�)
'strRobocopyStatus = objShell.Run("cmd /c robocopy.exe " & strLocalDir & " " & strRemoteDir & " " & " /E /COPYALL  /R:2 /W:5 /FFT /TEE /NDL /NP /LOG:" & strLogFile ,7,True)
'�~���[�o�b�N�A�b�v�̏ꍇ
strRobocopyStatus = objShell.Run("cmd /c robocopy.exe " & strLocalDir & " " & strRemoteDir & " " & " /LOG:" & strLogFile & " /COPY:DAT",7,True)

Dim strMsg , strStatus ,iRobocopyStatus , strInfo
iRobocopyStatus = CInt(strRobocopyStatus)
strStatus = "------------------------------" & vbCrLf  & "Status: " & vbCrLf

strInfo = "�o�b�N�A�b�v��:" & strLocalDir & vbCrLf & "�o�b�N�A�b�v��:"& strRemoteDir & vbCrLf

'robocopy����̖߂�l��_���ς�����ďڍ׃��b�Z�[�W����
If (iRobocopyStatus AND 16) = 16 Then
    strStatus = strStatus & "***FATAL ERROR***"
    strMsg = strMsg & "Code16:�d��ȃG���[�BRobocopy�́C�t�@�C�����ЂƂ��R�s�[���܂���ł����B����́C�g�p�@�̊ԈႢ���]�����܂��͓]����f�B���N�g���ւ̃A�N�Z�X�����s�\���Ȃ��߂ɔ�������G���[�ł��B" & vbCrLf 
End If
If (iRobocopyStatus AND 8) = 8 Then
    strStatus = strStatus & " FAIL "
    strMsg = strMsg & "Code8:�������̃t�@�C���܂��̓f�B���N�g�����R�s�[�ł��܂���ł����i�R�s�[�G���[���������C���g���C���������܂����j�B�ȍ~�̃G���[���`�F�b�N���Ă��������B" & vbCrLf 
End If
If (iRobocopyStatus AND 4) = 4 Then
    strStatus = strStatus & " MISM "
    strMsg = strMsg & "Code4:������Mismatched(�]�����̃t�@�C���Ɠ����̃f�B���N�g�����]����ɂ���,���邢�͂��̋t)�t�@�C���܂��̓f�B���N�g�����݂���܂����B���O�𒲂ׂĂ��������B�����|�����K�v�ł��B" & vbCrLf 
End If
If (iRobocopyStatus AND 2) = 2 Then
    strStatus = strStatus & " XTRA "
    strMsg = strMsg & "Code2:��������Extra(�]�����ɑ��݂��Ȃ��̂ɓ]����ɑ��݂���t�@�C��)�t�@�C���܂��̓f�B���N�g�����݂���܂����B�o�̓��O�𒲂ׂĂ��������B�|�����K�v��������܂���B" & vbCrLf 
End If
If (iRobocopyStatus AND 1) = 1 Then
    strStatus = strStatus & " COPY "
    strMsg = strMsg & "Code1:��ȏ�̃t�@�C�����C���܂��R�s�[����܂����i�܂�V�����t�@�C���͓]����ɓ͂��܂����j�B" & vbCrLf 
End If
If iRobocopyStatus = 0 Then
    strMsg = strMsg & "Code0:�G���[�͔��������C�R�s�[������܂���ł����B�]�����Ɠ]����f�B���N�g���c���[�́C���S�ɓ������Ă��܂��B" & vbCrLf 
End If

strStatus = strStatus & vbCrLf  & "---------------------------------"

'robocopy����̖߂�l�ɂ���ď�����U�蕪����8�ȏオ�v���I�ȃG���[
If iRobocopyStatus >= 8 Then
    strSubject = "[Robocopy_BackupError][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName
    SendMail strSubject , "Host:" & objNetWork.ComputerName & "�̃o�b�N�A�b�v�����ɒv���I�ȃG���[���������܂����B���O�t�@�C�����m�F���Ă��������B" & vbCrLf & _
    vbCrLf & strStatus & vbCrLf &  strInfo & vbCrLf & vbCrLf & "�ڍ׏��:" & vbCrLf & strMsg
Else
    strSubject = "[Robocopy_BackupLog][" & CStr(Now()) & "]" & objNetWork.UserName & "@" & objNetWork.ComputerName
    SendMail strSubject , "Host:" & objNetWork.ComputerName & "�̃o�b�N�A�b�v��������ɏI�����܂����B" & vbCrLf & _
    vbCrLf & strStatus & vbCrLf &  strInfo & vbCrLf & vbCrLf & "�ڍ׏��:" & vbCrLf & strMsg
End If

Set objNetWork = Nothing

'�����������������ł���(MS�̃h�L�������g���)
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
 
 
'���[�����M����
Function SendMail2(strSubject , strError)
	print strSubject
	print strError
End Function
Function SendMail(strSubject , strError)
    Dim strMailAddrTo , strSMTPServ , strMailAddrFrom , iSMTPPort
    '# ���[���A�h���X��ݒ�
    strMailAddrTo = "tytakayanagi@nissho-ele.co.jp"
    strSMTPServ = "172.25.34.1"
    strMailAddrFrom = "grandit-skc-support@nissho-ele.co.jp"
    iSMTPPort = 25

    Dim objMail
    Set objMail = CreateObject("CDO.Message")
    '���[���𑗐M
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