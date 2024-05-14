#One Drive scripted fix
Read-Host -Prompt "User with OneDrive to repair must be logged in. Press Enter to continue. ctrl+c to exit"
#set org
$Org = "Contoso"
# check for admin rights
if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Run as Admin
    # Get logged in username
    $LoggedOnUser = (qwinsta /SERVER:$env:COMPUTERNAME) -replace '\s{2,22}', ',' | ConvertFrom-Csv | Where-Object { $_ -like "*Acti*" } | Select-Object -ExpandProperty Username
    # Get SID of user via AD
    #$adSID = (Get-ADUser $LoggedOnUser).SID.Value
    # Get SID of user via registry
    $rSID=(Get-ItemProperty -Path registry::"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -Name ProfileImagePath | Where-Object { $_.ProfileImagePath -like "*$LoggedOnUser" }).PSChildName
    # Registry path for OneDrive
    #$reg = "HKEY_USERS\$adSID\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $reg = "HKEY_USERS\$rSID\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    # Create OneDrive path prefix
    $odPathPrefix = "C:\Users\$LoggedOnUser\OneDrive - $org"
}
else {
    # Run as user
    # Get user name
    $sUser = $(whoami).split('\')[1]
    # Get user profile path
    $reg = "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    # Create OneDrive path prefix
    $odPathPrefix = "C:\Users\$sUser\OneDrive - $org"
}

# properties to check
$props = 'Desktop', 'Favorites', "My Music", "My Pictures", "My Video", "Personal", "{35286A68-3C57-41A1-BBB1-0EAE73D76C95}", "{0DDD015D-B06C-45D5-8C4C-F59713854639}", "{A0C69A99-21C8-4671-8703-7934162FCF1D}", "{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}", "{F42EE2D3-909F-4907-8871-4C22FC0BF756}", "{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}" 
ForEach ($prop in $props) {
    # Get OneDrive paths from registry
    $testVal = (Get-ItemProperty "registry::$reg" -Name $prop).$prop
    # Check if path is OneDrive path
    if ($testVal -like "$odPathPrefix*") {
        # Path already OneDrive Format
        Write-Host $prop "is already in OneDrive format"
        continue
    }
    else {
        Write-Host $prop "is not OneDrive format" -BackgroundColor Red -ForegroundColor White
        # New path
        $newPath = "$odPathPrefix\$prop"
        Write-Host "Setting" $prop "to" $newPath
        # Set new path
        Set-ItemProperty registry::"$reg" -Name $([string]$prop) -Value $newPath -Type ExpandString
        # Check if path was set
        if ($(Get-ItemProperty "registry::$reg" -Name $prop).$prop -eq $newPath) {
            # Path set successfully
            Write-Host "Successfully set $prop to $newPath" -BackgroundColor Green -ForegroundColor White
        }
        else {
            # Path failed to set
            Write-Host "Failed to set $prop to $newPath" -BackgroundColor Red -ForegroundColor White
        }
    }
}   