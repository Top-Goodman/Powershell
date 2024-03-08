[CmdletBinding()]
Param (
    [Parameter(Mandatory = $true)]
    [array]$CompList = @('',''),
    [Parameter(Mandatory = $true)]
    [int]$addSpace = 5,
    [Parameter(Mandatory = $true)]
    [switch]$Force = $false,
    [Parameter(Mandatory = $true)]
    [switch]$Verbose = $false,
    [Parameter(Mandatory = $true)]
    [array]$servers = @('','')
)
### Variables
[int32]$addCap=10
[int32]$maxDiskSizeGB=120

### DO NOT EDIT BELOW ###

# Limit the amount that the script can add at once
if ($addSpace -gt $addCap) {
    $null = Read-Host "Can not Add More than 10GB to a VM. Press ENTER to continue" | Out-Null
    break
}
# Install the module so script can interact with VMWare api, only if not already available to system
$Module = "VMware.PowerCLI"
if (Get-Module -ListAvailable -Name $Module) {
    Write-Host "VMWare Power CommandLine already Installed"
}
else {
    try {
        Install-Module $Module -Scope CurrentUser -Force -AllowClobber
        Write-Host "VMWare Power CommandLine Installed"
    }
    Catch {
        Write-Host "$($_.Exception.Message)"
        Write-Host "Issue with VMWare PowerCLI Dependency. Module not able to be installed"
        #exit 1
    }
}
# Import the module so script can interact with VMWare api, only if not already loaded on system
If ($Module -in $(Get-Module).Name) {
    Write-Host "$Module already Imported"
}
Else {
    Try {
        Import-Module $Module -Force -ErrorAction Stop
        Write-Host "$Module successfully Imported"
    }
    Catch {
        Write-Host "$($_.Exception.Message)"
        Write-Host "Issue with $Module. Module not able to be imported"
        #exit 1
    }
}

# Connect to servers
Connect-VIServer $servers
$CompList | ForEach-Object {
    try {
        Write-Host "Getting VM informaiton"
        $VMx = $_
        # Find VM
        $vm = Get-VM | Where-Object { $_.Name -eq $VMx }
        Write-Host $vm
        # Get VDisk indo
        $HD = Get-HardDisk -VM $vm.Name -Name "Hard disk 1"
        Write-Host $HD
        $targetCap = $HD.CapacityGB + $addSpace
        #   $NewCap = [decimal]::round($HD.CapacityGB + 5, $maxDiskSizeGB) 
       
       
        # Notify that target drive size too many big 
        if (!($Force)) {
        # Determine if new disk is higher than scripted disk limit (120GB)
            $NewCap = [Math]::Min($HD.CapacityGB + $addSpace, $maxDiskSizeGB)
            if ($targetCap -gt $maxDiskSizeGB) {
                $null = Read-Host "New disk must be less than $($maxDiskSizeGB)GB. Selection would make new disk $($targetCap)GB. Press ENTER to continue" | Out-Null
                break
            }
                    Write-Host "New-Cap is $newcap"
        }
if ($Verbose) {}
else {}
       # $HD | Set-HardDisk -CapacityGB $NewCap -Confirm:$false
        Write-Host "VDisk added"
    }
    catch {
        Write-Host "$($_.Exception.Message)"
        Write-Host "Issue with Adding to disk. Module not able to be installed"        
    }
}
Write-Host "Now script will expand Windows Drive Partition.....!!!!"

$CompList | ForEach-Object {
    $global:vmName = $_
    Try {
        Write-Host "Creating Diskpart.txt"
        # Create diskpart text to variable with drive letter of taget VM
        $global:diskpart = @"
rescan
select volume $((Invoke-VMScript -VM $global:vmName -ScriptText `$env:systemdrive).ScriptOutput.replace("`n","").replace("`r","").replace(':',""))
extend
exit
"@
        Write-Host "Writing Diskpart.txt to \\$vmName\C$\Windows\temp\diskpart.txt"
        # Write variable to target VM in ascii (UTF-8) encoding so diskpart can reader
        $global:diskpart | Out-File "\\$vmName\C$\Windows\temp\diskpart.txt" -Force -Encoding ascii -Confirm:$false
        Write-Host "Expanding Disk"
        # Tell vm to run diskpart and expand disk
       # Invoke-VMScript -VM $global:vmName -ScriptText "C:\windows\system32\diskpart.exe /s c:\Windows\temp\diskpart.txt" -ScriptType BAT
    }
    catch {
        Write-Host "$($_.Exception.Message)"
        Write-Host "Issue with expanding disk."  
    }
}