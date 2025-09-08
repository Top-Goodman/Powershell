param (
    [Parameter(Mandatory = $true)]
    [string]$Path2Map,

    [switch]$NoCredentials
)

# Get drive letters C: to Z:
$Letters = 67..90 | ForEach-Object { [char]$_ + ":" }

# Get currently mapped drives
$Mapped = Get-WmiObject Win32_LogicalDisk | Select-Object DeviceID, ProviderName
$InUse = $Mapped.DeviceID
$NetProvider = $Mapped.ProviderName

# Check if the network path is already mapped
foreach ($nPath in $NetProvider) {
    if ($nPath -eq $Path2Map) {
        Write-Output "Share already mapped... Exiting"
        return
    }
}

Write-Output "$Path2Map not found. Continue to Mapping..."

# Find an available drive letter
$FreeLetter = $null
foreach ($Letter in $Letters) {
    if ($InUse -contains $Letter) {
        Write-Output "$Letter is in use"
    } else {
        Write-Output "$Letter is available"
        $FreeLetter = $Letter
        break
    }
}

# Map the drive
if ($FreeLetter) {
    try {
        if ($NoCredentials) {
            New-PSDrive -Name $FreeLetter.TrimEnd(':') -PSProvider FileSystem -Root $Path2Map -Persist
        } else {
            $cred = Get-Credential
            New-PSDrive -Name $FreeLetter.TrimEnd(':') -PSProvider FileSystem -Root $Path2Map -Credential $cred -Persist
        }
        Write-Output "Drive mapped to $FreeLetter"
    } catch {
        Write-Error "Failed to map drive: $_"
    }
} else {
    Write-Error "No available drive letters found."
}
