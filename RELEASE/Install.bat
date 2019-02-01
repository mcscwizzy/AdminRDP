REM Make new directories and copy files
md "C:\Program Files\AdminRDP"
if exist "C:\Program Files\AdminRDP\SavedServers\ServerList.xml" (
	xcopy "%~dp0*.*" "C:\Program Files\AdminRDP" /y
	xcopy "%~dp0AdminRDP.lnk" "%USERPROFILE%\Desktop" /y
) else (
	xcopy "%~dp0*.*" "C:\Program Files\AdminRDP" /s /y
	xcopy "%~dp0AdminRDP.lnk" "%USERPROFILE%\Desktop" /y
)

"%USERPROFILE%\Desktop\AdminRDP.lnk"


