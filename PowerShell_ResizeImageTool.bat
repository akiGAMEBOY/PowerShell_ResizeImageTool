@ECHO OFF
@REM #################################################################################
@REM # �������@�bResizeImageTool�i�N���p�o�b�`�j
@REM # �@�\�@�@�bPowerShell�N���p�̃o�b�`
@REM #--------------------------------------------------------------------------------
@REM # �@�@�@�@�b-
@REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  ResizeImageTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.
SET RETURNCODE=0
@REM PowerShell Core �C���X�g�[���m�F
WHERE /Q pwsh
IF %ERRORLEVEL% == 0 (
    @REM PowerShell Core �Ŏ��s����ꍇ
    pwsh -NoProfile -ExecutionPolicy Unrestricted -File .\source\powershell\Main.ps1
) ELSE (
    @REM PowerShell 5.1  �Ŏ��s����ꍇ
    powershell -NoProfile -ExecutionPolicy Unrestricted -File .\source\powershell\Main.ps1
)

SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO �������I�����܂����B
ECHO �����ꂩ�̃L�[�������ƃE�B���h�E�����܂��B
PAUSE > NUL
EXIT %RETURNCODE%