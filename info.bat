@echo off
:filename for recording information
set fname=pcinfo.txt
:PC name
Echo pcname: %computername% >>%fname%
:IP address of the computer by name
FOR /F "usebackq tokens=2 delims=[]" %%i IN (`ping %Computername% -n 1 -4`) DO if not "%%i"=="" Set ip=%%i
Echo IP %computername%: %ip% >>%fname%
:active username
Echo username: %username%  >>%fname%
:laptop model
set cmd=wmic computersystem get model
for /f "skip=1 delims=" %%Z in ('%cmd%') do (
    set _pn=%%Z
	GOTO BREAK1
)
:BREAK1
echo CS Model: %_pn% >>%fname%
:CPU
set cmd=wmic cpu get name
for /f "skip=1 delims=" %%Z in ('%cmd%') do (
    set _cpu=%%Z
	GOTO BREAK1
)
:BREAK1
echo CPU: %_cpu% >>%fname%
:motherboard
set cmd=wmic baseboard get product
for /f "skip=1 delims=" %%Z in ('%cmd%') do (
    set _mb=%%Z
    GOTO BREAK2
)
:BREAK2
echo MB: %_mb% >>%fname%
:RAM
SETLOCAL ENABLEDELAYEDEXPANSION
set mmr=0
for /f "skip=1 delims=" %%i in ('WMIC MemoryChip get BankLabel^,DeviceLocator^,PartNumber^,Speed^,Capacity') do (
for /f "tokens=1-5 delims=" %%A in ("%%i") do (
set BnkLbl=%%A
set /a mmr=!mmr!+1
set BnkLbl=!BnkLbl:BANK 22=DDR2 FB-DIMM!
set BnkLbl=!BnkLbl:BANK 21=DDR2!
set BnkLbl=!BnkLbl:BANK 24=DDR3!
set BnkLbl=!BnkLbl:BANK 0=DDR4!
echo Memory !mmr!: !BnkLbl! >>%fname%
))
:hard disks
SETLOCAL ENABLEDELAYEDEXPANSION
set mmr=0
for /f "skip=1 delims=" %%i in ('wmic diskdrive get model^,size') do (
for /f "tokens=1-2 delims=" %%A in ("%%i") do (
set HDDLbl=%%A
set /a mmr=!mmr!+1
echo HDD !mmr!: !HDDLbl! >>%fname%
))