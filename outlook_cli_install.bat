@ECHO OFF
SETLOCAL ENABLEEXTENSIONS

:checkAdmin
fsutil dirty query %systemdrive% >nul
IF %errorLevel% == 0 (
    ECHO This script can only be run as the current user.
    ECHO Do not run as administrator.
    PAUSE
    EXIT
)

:checkArch
IF EXIST "%PROGRAMFILES(X86)%" (
    SET "VSTOINSTPATH=C:\Program Files (x86)\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe"
) ELSE (
    SET "VSTOINSTPATH=C:\Program Files\Common Files\microsoft shared\VSTO\10.0\VSTOInstaller.exe"
)

:aTest
echo This is a test

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

SET "G2MPATH=https://builds.citrixonlinecdn.com/builds/calendarintegration/outlook/G2M/GoToMeetingOutlookCalendarPlugin.vsto"


IF %PROD% EQU 1 (
    SET "PRODPATH=https://builds.citrixonlinecdn.com/builds/calendarintegration/outlook/G2M/GoToMeetingOutlookCalendarPlugin.vsto"
	) ELSE (
	    IF %PROD% EQU 2 (
		    SET "PRODPATH=https://builds.citrixonlinecdn.com/builds/calendarintegration/outlook/openvoice/OpenVoiceForOutlook.vsto"
		) ELSE (
		    CLS
			ECHO Invalid selection. Please try again.
			PAUSE
			GOTO :getProd
		)
	)
)

"%VSTOINSTPATH%" /install %PRODPATH%
