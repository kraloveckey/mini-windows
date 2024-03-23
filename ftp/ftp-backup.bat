set @pathBackup=C:\BackupSQL\FOLDER_NAME-%time:~0,2%-%time:~3,2%-%time:~6,2%_%date:~-10,2%-%date:~-7,2%-%date:~-4,4%.zip
"C:\Program Files\7-Zip\7z.exe" a -tzip %@pathBackup% C:\BackupSQL\FOLDER_NAME

ncftpput -R -v -u "USERNAME" -p "PASSWORD" IP_HOST /FOLDER_NAME %@pathBackup%