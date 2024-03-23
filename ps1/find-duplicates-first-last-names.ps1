# Hash table of GivenName/Surname combinations.
$Names = @{}

# Retrieve all users with GivenName and Surname.
$Users = Get-ADUser -Filter 'GivenName -Like "*" -And Surname -Like "*"' | Select GivenName, Surname, sAMAccountName

# Find all users with each unique GivenName/Surname combination.
ForEach ($User In $Users)
{
    $Name = $User.GivenName + "/" + $User.Surname
    $NTName = $User.sAMAccountName
    If ($Names.ContainsKey($Name))
    {
        # Duplicate. Add the sAMAccountName to the value, delimited by a semicolon.
        # Semicolons are not allowed in sAMAccountName.
        $Names[$Name] = $Names[$Name] + ";" + $NTName
    }
    Else {$Names.Add($Name, $NTName)}
}

# Output entries in the hash table where the value includes the ";" character, indicating duplicate accounts.
ForEach ($Key In $Names.Keys)
{
    If ($Names[$Key] -Like "*;*") {$Key + " : " +$Names[$Key]}
}