Write-Output "Generating passwords for the users in users.csv..."
$Users = Get-Content ./users.csv
$Result = @()

ForEach($user in $Users){
    $password = ((0..3 | foreach {(Get-Culture).TextInfo.ToTitleCase($(irm "http://randomword.setgetgo.com/get.php?len=$(get-random -min 5 -max 9)"))}) -join '') + '-' + (get-random 999)
    $Result += "$($user);$($password)"
}
$Result
$Result > passwords.csv