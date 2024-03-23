$users = Import-Csv -Path .\change-userdata.csv
foreach ($user in $users) {
    $First = $user.Name
    $Last = $user.Surname
    $Display = $user.Full
    Get-ADUser -Filter "mail -eq '$($user.Mail)'" -Properties * | Set-ADUser -givenName $First -displayname $Display -Surname $Last
    Get-ADUser -Filter "mail -eq '$($user.Mail)'" -Properties * | Rename-ADObject -NewName $Display
}