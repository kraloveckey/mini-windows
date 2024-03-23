# Replace 'YourOU' with the name of your Organizational Unit and 'old_domain.com' with the current domain
$ouName = "OU=Users,DC=dns,DC=com"
$oldDomain = "dns-old.com"
$newDomain = "dns-new.com"  # New domain to replace the existing one

# Get users from the specified OU
$users = Get-ADUser -Filter * -SearchBase "$ouName" -Properties EmailAddress

# Loop through each user and change the domain part of their email address
foreach ($user in $users) {
    $currentEmail = $user.EmailAddress

    # Check if the email is not empty and contains the old domain
    if ($currentEmail -ne $null -and $currentEmail -like "*@$oldDomain") {
        $newEmail = $currentEmail -replace "@$oldDomain", "@$newDomain"
        
        # Update the user's email address
        Set-ADUser -Identity $user -EmailAddress $newEmail

        Write-Host "Changed email for $($user.SamAccountName) from $currentEmail to $newEmail"
    } else {
        Write-Host "Skipping $($user.SamAccountName) - Email not found or does not match the old domain"
    }
}