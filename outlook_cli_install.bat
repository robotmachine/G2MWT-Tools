@ECHO OFF
setlocal EnableExtensions EnableDelayedExpansion

TITLE .:. GoToMeeting and OpenVoice for Outlook Helper .:.

:checkAdmin
REM This command does not do anything useful
REM It only exists to check to see if we have admin rights
REM This command fails if run as a user and succeeds if run as admin
REM After running the command and redirecting all output to nowhere
REM we check the exit code to guess if we are running as admin or not.
fsutil dirty query %systemdrive% >nul
IF %errorLevel% EQU 0 (
    ECHO This script should not be run as administrator and it
    ECHO seems possible that you have done so.
    ECHO --
    ECHO Please type "0" to bypass if this is not correct or
    ECHO press enter to exit.
    SET "bypass=1"
    SET /p bypass=".:. "
    IF !BYPASS! EQU 0 (
        GOTO :checkArch
        ) ELSE (
            EXIT
        )
)

:checkRegistry
REG QUERY "HKCU\Software\Microsoft\Office\Outlook\Addins\GoToMeetingOutlookCalendarPlugin"
IF %errorLevel% EQU 0 (
    CLS
	ECHO GoToMeeting for Outlook might already be installed
	ECHO Please uninstall from Programs and Features before proceeding
	PAUSE)
	
REG QUERY "HKCU\Software\Microsoft\Office\Outlook\Addins\OpenVoiceForOutlook"
IF %errorLevel% EQU 0 (
	CLS
	ECHO OpenVoice for Outlook might already be installed
    ECHO Please uninstall from Programs and Features before proceeding
	PAUSE)

:checkArch
REM For lack of a more sensible method, we just look to see if the Program Files (x86) directory exits
REM to determine if we're on a 32 or 64 bit OS
IF EXIST "%PROGRAMFILES(X86)%" (
    SET "VSTOINSTPATH=C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe"
) ELSE (
    SET "VSTOINSTPATH=C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe"
)

:checkVSTools
IF NOT EXIST "%VSTOINSTPATH%" (
        ECHO Visual Studio 2010 Tools for Outlook is not installed correctly.
        ECHO Please go to https://bit.ly/vstools to install.
        PAUSE
		EXIT
)

:getProd
CLS
ECHO Select Product
ECHO 1) GoToMeeting for Outlook
ECHO 2) OpenVoice for Outlook
SET /p prod="Enter Selection: "

IF %PROD% EQU 1 (
    SET "PRODPATH=https://builds.getgocdn.com/builds/calendarintegration/outlook/G2M/GoToMeetingOutlookCalendarPlugin.vsto"
	) ELSE (
	    IF %PROD% EQU 2 (
		    SET "PRODPATH=https://builds.getgocdn.com/builds/calendarintegration/outlook/openvoice/OpenVoiceForOutlook.vsto"
		) ELSE (
		    CLS
			ECHO Invalid selection. Please try again.
			PAUSE
			GOTO :getProd
		)
	)
)

"%VSTOINSTPATH%" /install %PRODPATH%
