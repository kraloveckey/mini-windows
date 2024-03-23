@echo off
	set Log=C:\users\%username%\desktop\Log.txt

:Start
echo.
	set /p action=Full-1^|RestoreHealth-2^|Sfc-3^:
	if /i "%action%"=="1" (goto Full) 
	if /i "%action%"=="2" (goto RestoreHealth) 
	if /i "%action%"=="3" (goto Sfc) 
exit 

:Full
	echo Start %time% 
	echo "Full Scanning"
	::echo Create Log.txt
	echo "DISM /Online /Cleanup-Image /CheckHealth"
	DISM /Online /Cleanup-Image /CheckHealth 

	echo "DISM /Online /Cleanup-Image /ScanHealth"
	DISM /Online /Cleanup-Image /ScanHealth 

	echo "DISM /Online /Cleanup-Image /RestoreHealth"
	DISM /Online /Cleanup-Image /RestoreHealth 
 
	echo "sfc /scannow"
	sfc /scannow  
goto Done

:RestoreHealth
	echo Start %time% 
	echo "RestoreHealth"
	::echo Create Log.txt
	echo "DISM /Online /Cleanup-Image /RestoreHealth"
	DISM /Online /Cleanup-Image /RestoreHealth >> "%Log%" 
		find "Error" %Log% 
		if %errorlevel% equ 1 goto Notfound
goto Done

:Sfc
	echo Start %time% 
	echo "Sfc Scan"
	::echo Create Log.txt
	sfc /scannow  
goto Done

:Notfound
	echo Good
	goto Done


:Done
	echo Stop %time%
	pause
goto Action

:Action
	echo You want to continue or exit the program?
	set /p action=TO CONTINUE OR EXIT? [Y\N]:
	if /i "%action%"=="Y" (goto Start) else exit