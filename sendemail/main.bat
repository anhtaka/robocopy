@echo off

rem ���t���擾
set YYYYMMDD=%DATE:/=%

rem �R�s�[���A�R�s�[��A���O�t�@�C�����̐ݒ�t�@�C���Ăяo��
FOR /F "tokens=1,2,3 delims=/" %%A IN (robocopy_bkup_conf.txt) DO (
   rem VB�X�N���v�g�����s %%A:�o�b�N�A�b�v�� %%B:�o�b�N�A�b�v�� %%C:���O�t�@�C����
   cscript "robocopy.vbs" %%A %%B %%ClogrobocopyLog_%YYYYMMDD%.txt
)

rem �Â����O�t�@�C���폜
rem cscript "LogOldFile_Del.vbs"