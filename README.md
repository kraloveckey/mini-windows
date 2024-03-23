# Mini-Windows Projects

## ADInventory

[`ADInventory.ps1`](./ADInventory/ADInventory.ps1) - script to inventory computers in Active Directory. After launching it polls the specified computers in [**OU.txt**](./ADInventory/OU.txt) and sends to mail an archive with files for each OU from [**OU.txt**](./ADInventory/OU.txt) (**OU-Computers.csv** and **OU-Software.csv** - specifying the name of the OU).

[`OU.txt`](./ADInventory/OU.txt) - this file must contain the full paths of the OU with Active Directory computers. This file must be located in the same directory when starting [**ADInventory.ps1**](ADInventory.ps1).

[`.env`](./ADInventory/.env) - the file in which contain the password for SMTP login (for **AUTH_USER@**).

### Description of the files that are sent to the mail in the archive after [ADInventory.ps1](./ADInventory/ADInventory.ps1) execution

| Name      |  Description |
| ----------- |  ----------- |
| `OU-Computers.csv` | Contains information on computers: Name, When online, OU, OS, OS Version, IP, CPU, Frequency - MHz, Number of cores, RAM capacity - MB, Drive capacity - GB, Drive models. |
| `OU-Software.csv` | Contains information on software, the file contains a list of all found software and check marks against computers where it is installed. |

## ftp

[`ftp-backup.bat`](ftp/ftp-backup.bat) - script for backup **FOLDER_NAME** and upload archive to FTP Server: **IP_HOST**.

[`ftp-reload.bat`](ftp/ftp-reload.bat) - script for reload FTP Service.

[`rollback-up.bat`](ftp/rollback-up.bat) - script for move beta project to prod and prod to old folder and rollback its changes.

## lnk

[`Lock.lnk`](lnk/Lock.lnk) - shortcut for lock Windows OS.

[`Restart.lnk`](lnk/Restart.lnk) - shortcut for restart Windows OS.

[`Shutdown.lnk`](lnk/Shutdown.lnk) - shortcut for shutdown Windows OS.

## ps1

[`change-email.ps1`](ps1/change-email.ps1) - script to update the user's email address, where **@old_domain.com** replacing with **@new_domain.com**. 

[`change-userdata-by-mail.ps1`](ps1/change-userdata-by-mail.ps1) - script to change user data (Full Name, First Name, Surname) by email address from **change-userdata.csv** file.
   - [`change-userdata.csv`](ps1/change-userdata.csv) - file contains user data (Full Name, First Name, Surname) for change by email address.

[`find-duplicates-first-last-names.ps1`](ps1/find-duplicates-first-last-names.ps1) - script to finding duplicates user by **GivenName/Surname** combinations with adding the **sAMAccountName** to the each value.

## Other

[`check.bat`](check.bat) - script for scan with `DISM` and `sfc` Windows OS.

[`Disable SMBv1.reg`](Disable%20SMBv1.reg) - reg key for disable SMBv1.

[`info.bat`](info.bat) - script for get full information from Windows device and record its to **pcinfo.txt** file.

[`Wallpapers.deskthemepack`](Wallpapers.deskthemepack) - `Surreal Territory` themepack of photographic illustration wallpapers.