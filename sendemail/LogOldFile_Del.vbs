'log�t�@�C���̊i�[����Ă���t�H���_�[�̃p�X
Const cFolderPath = "D:\sscwork\00.�f�[�^�ڍs\20180414_�{�Ԉڍs\robocopy\log"
'�ȉ��̕����񂪐擪�ɂ���t�@�C�����폜����
Dim strLogShareFile
strLogShareFile = "robocopyLog_"
'�ȉ��̓����o���Ă�����̂��폜�ΏۂƂ���
Dim iDate
iDate=120

Dim fso, fl
Dim tod, logDate, buf, diff
Set fso = CreateObject("Scripting.FileSystemObject")

tod = CDate(Split(Now, " ")(0))
For Each fl In fso.GetFolder(cFolderPath).Files
    ' strLogShareFile���擪�ɂ��t�@�C���̂ݍ폜�ΏۂƂ���
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