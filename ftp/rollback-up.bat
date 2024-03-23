@echo off
	set Log=C:\users\%username%\Desktop\Log.txt
	set @pathProd=C:\users\%username%\Desktop\Prod
	set @pathOld=C:\users\%username%\Desktop\Old
	set @pathBeta=C:\users\%username%\Desktop\Beta

:Start
echo.
	set /p action=UpProd-1^|Rollback-2^:
	if /i "%action%"=="1" (goto UpProd) 
	if /i "%action%"=="2" (goto Rollback) 
exit 

:UpProd
	echo Start %time% 

	set @nameOld=NAME_%time:~0,2%-%time:~3,2%-%time:~6,2%_%date:~-10,2%-%date:~-7,2%-%date:~-4,4%
	echo Moving from %@pathProd% to %@pathOld%\%@nameOld%
	md %@pathOld%\%@nameOld%
	Powershell.exe Move-Item -Path %@pathProd%\* -Destination %@pathOld%\%@nameOld%
	
	echo Moving from %@pathBeta% to %@pathProd%
	Powershell.exe Move-Item -Path %@pathBeta%\* -Destination %@pathProd%
goto Done

:Rollback
	echo Start %time% 
	echo Rollback...
	echo Cleaning %@pathBeta%
	Powershell.exe Remove-Item %@pathBeta%\* -Recurse

	echo Moving from %@pathProd% to %@pathBeta%
	Powershell.exe Move-Item -Path %@pathProd%\* -Destination %@pathBeta%

	echo Search last modified folder...
	for /f "tokens=*" %%A in ('dir %@pathOld% /AD /O-D /B') do (set recent=%%A& goto exit)

	:exit
	echo Last modified folder - %recent%
	echo Moving from %@pathOld%\%recent% to %@pathProd%
	Powershell.exe Move-Item -Path %@pathOld%\%recent%\* -Destination %@pathProd%

	echo Deleting folder - %@pathOld%\%recent%
	Powershell.exe Remove-Item %@pathOld%\%recent%
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